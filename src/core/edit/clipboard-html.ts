/**
 * 解析从 Excel/WPS 复制的剪贴板 HTML(text/html)→ 值 + 样式 + 合并 + 图片。
 * 与 WPS 粘贴一致地还原字体/颜色/填充/边框/对齐/合并;值优先取 Excel 的 x:num/x:fmla(原始值),
 * 否则取单元格文本(交给 setCellValue 推断)。图片仅 data-uri <img>(区域复制一般拿不到内嵌图,见 README)。
 *
 * 用浏览器原生 DOMParser(框架无关;core 仍可 headless,缺 DOMParser 时返 null 回退 TSV)。
 */
import { unzipSync } from 'fflate'
import type { BorderEdge, BorderStyle, CellStyleOverride, Font, MergeRange } from '../model/types'
import type { CellValue } from '../model/data-access'
import { toHex6 } from '../format/color'
import { b64ToBytes, bytesToB64 } from './clipboard-snapshot'

export interface ParsedClipboard {
  /** 二维值(原始串/数字/公式,交 setCellValue 推断);稠密对齐 */
  values: CellValue[][]
  /** 逐格样式覆盖(只含解析出的字段) */
  styles: { row: number; col: number; patch: CellStyleOverride }[]
  /** 合并区(rowspan/colspan>1) */
  merges: MergeRange[]
  /** data-uri 图片(row,col,dataUrl) */
  images: { row: number; col: number; dataUrl: string }[]
  /** 列宽(相对列 index → px;来自 <col width>);稀疏 */
  colWidths: number[]
  /** 行高(相对行 index → px;来自 <tr height>);稀疏 */
  rowHeights: number[]
}

/** css border-style → 我们的 BorderStyle(近似) */
function mapBorderStyle(css: string): BorderStyle | null {
  switch (css) {
    case 'solid':
      return 'thin'
    case 'dotted':
      return 'dotted'
    case 'dashed':
      return 'dashed'
    case 'double':
      return 'double'
    case 'none':
    case 'hidden':
      return null
    default:
      return 'thin'
  }
}

/** 解析一条 css 边框简写(如 "1px solid #d4d4d4")→ BorderEdge(无有效线返 undefined) */
function parseBorderEdge(spec: string): BorderEdge | undefined {
  if (!spec) return undefined
  const styleMatch = /\b(solid|dotted|dashed|double|none|hidden)\b/i.exec(spec)
  const style = styleMatch ? mapBorderStyle(styleMatch[1].toLowerCase()) : null
  if (!style) return undefined
  const colorMatch = /#[0-9a-f]{3,8}\b|rgba?\([^)]+\)/i.exec(spec)
  return { style, color: (colorMatch && toHex6(colorMatch[0])) || '#000000' }
}

/**
 * 收集文档所有 <style> 块里的「类名 → 声明串」(Excel/WPS 复制把格式放这儿,如 `.xl65{border:...;background:...}`),
 * 外加裸 `td` 元素选择器的默认声明(WPS 把"所有单元格默认"——如 `vertical-align:middle`、`white-space:nowrap`——
 * 放在 `td{...}` 上,各 `.etN` 类只覆盖要改的;不收这层就丢掉默认垂直居中等)。
 * 类:只取选择器末尾的 class token(`.xl65` / `td.xl65` / `.xl65,.xl66` 都认);同类多条按出现序拼接(后者覆盖)。
 */
function parseClassStyles(doc: Document): { classes: Map<string, string>; tdDefault: string } {
  const map = new Map<string, string>()
  let tdDefault = ''
  doc.querySelectorAll('style').forEach((styleEl) => {
    // Office 把 CSS 整段包在 <!-- --> 里(.font0 前直接跟 <!--),先剥掉注释壳;再去 /* */ 注释
    const css = (styleEl.textContent || '').replace(/<!--|-->/g, '').replace(/\/\*[\s\S]*?\*\//g, '')
    const ruleRe = /([^{}]+)\{([^{}]*)\}/g
    let m: RegExpExecArray | null
    while ((m = ruleRe.exec(css))) {
      const decls = m[2].trim()
      if (!decls) continue
      for (const sel of m[1].split(',')) {
        const s = sel.trim()
        const cls = /\.([A-Za-z0-9_-]+)\s*$/.exec(s)?.[1]
        if (cls) map.set(cls, map.has(cls) ? `${map.get(cls)};${decls}` : decls)
        else if (/(^|\s)td$/i.test(s)) tdDefault = tdDefault ? `${tdDefault};${decls}` : decls // 裸 td 默认层
      }
    }
  })
  return { classes: map, tdDefault }
}

/**
 * 合并单格样式声明,按 CSS 优先级从低到高拼:**td 元素默认 < 类规则 < 内联 style=**,写进 td.style 供 cssToStyleOverride 读取,
 * 并返回**合并后的原始声明串**(供解析 mso-number-format 等 CSSOM 会丢弃的私有属性)。
 */
function rawCssOf(td: HTMLElement, classStyles: Map<string, string>, tdDefault: string): string {
  const classNames = (td.getAttribute('class') || '').split(/\s+/).filter(Boolean)
  const classCss = classStyles.size ? classNames.map((c) => classStyles.get(c)).filter(Boolean).join(';') : ''
  const inline = td.getAttribute('style') || ''
  const combined = [tdDefault, classCss, inline].filter(Boolean).join(';')
  if (tdDefault || classCss) td.style.cssText = combined // 让 cssToStyleOverride 读到默认层+类里的标准属性(边框/底色/字体/对齐/垂直居中)
  return combined
}

/** mso-number-format 的值是 CSS 转义的 Excel 格式码(\0022→" \#→# \;→; \\(→\( …),解回真实格式码。 */
export function unescapeMsoNumFmt(v: string): string {
  return v
    .trim()
    .replace(/^"|"$/g, '') // 去外层引号
    .replace(/\\([0-9a-fA-F]{4})/g, (_, h) => String.fromCharCode(parseInt(h, 16))) // \0022 → "
    .replace(/\\(.)/g, '$1') // \# → #、\\( → \( 等单字符转义
}

/**
 * 从合并 CSS 串里解析 mso-number-format → Excel 格式码(值可能含转义的 \; ,故按引号串/到分号匹配)。
 * **取最后一条**:合并串是 `td默认;类;内联`,按 CSS 层叠后写的覆盖——裸 `td` 默认常带 `mso-number-format:General`,
 * 若取第一条会被它的 General 顶掉、丢掉后面类里的真实日期/货币格式码。
 */
export function parseMsoNumberFormat(css: string): string | undefined {
  const re = /mso-number-format:\s*("(?:\\.|[^"\\])*"|[^;]+)/gi
  let last: string | undefined
  let m: RegExpExecArray | null
  while ((m = re.exec(css))) last = m[1]
  if (last === undefined) return undefined
  const code = unescapeMsoNumFmt(last)
  return code && !/^general$/i.test(code) ? code : undefined
}

/**
 * WPS 区域复制把内嵌图放在 VML `<v:shape o:gfxdata="base64">`(在 `<!--[if gte vml 1]>…<![endif]-->` 注释里);
 * 那段 base64 是个 zip,内含 `media/imageN.png`。这里从 td 的注释节点取 o:gfxdata → 解 zip → 拿图 → data-uri。
 * (旁边的 `<img src="file:///…">` 是本地路径,浏览器读不了,只能靠这条。)
 */
function extractVmlImageDataUrl(td: Element): string | null {
  for (const node of Array.from(td.childNodes)) {
    if (node.nodeType !== 8) continue // 注释节点
    const data = (node as Comment).data
    const m = /o:gfxdata="([^"]+)"/.exec(data)
    if (!m) continue
    try {
      const files = unzipSync(b64ToBytes(m[1].replace(/\s+/g, '')))
      for (const [name, content] of Object.entries(files)) {
        const ext = /\.(png|jpe?g|gif|bmp)$/i.exec(name)?.[1]?.toLowerCase()
        if (ext && content.length) {
          const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`
          return `data:${mime};base64,${bytesToB64(content)}`
        }
      }
    } catch {
      /* 解 zip 失败 → 跳过 */
    }
  }
  return null
}

/** 单元格 css(td.style)→ CellStyleOverride;无可解析返 null */
function cssToStyleOverride(el: HTMLElement): CellStyleOverride | null {
  const st = el.style
  const patch: CellStyleOverride = {}
  const font: Partial<Font> = {}

  const fw = st.fontWeight
  if (fw === 'bold' || fw === 'bolder' || (fw && Number(fw) >= 600)) font.bold = true
  if (el.querySelector('b,strong')) font.bold = true
  if (st.fontStyle === 'italic' || el.querySelector('i,em')) font.italic = true
  const deco = (st.textDecorationLine || st.textDecoration || '') + (el.querySelector('u') ? ' underline' : '')
  if (/underline/.test(deco)) font.underline = true
  if (/line-through/.test(deco)) font.strike = true
  const color = toHex6(st.color)
  if (color) font.color = color
  if (st.fontFamily) font.name = st.fontFamily.split(',')[0].trim().replace(/['"]/g, '')
  if (st.fontSize) {
    const v = parseFloat(st.fontSize)
    if (v > 0) font.size = /pt\b/i.test(st.fontSize) ? Math.round(v) : Math.round((v * 72) / 96) // pt 原样;px → pt
  }
  if (Object.keys(font).length) patch.font = font

  const bg = toHex6(st.backgroundColor)
  if (bg) patch.fill = { type: 'solid', fgColor: bg }

  if (st.textAlign === 'left' || st.textAlign === 'center' || st.textAlign === 'right') patch.hAlign = st.textAlign
  const va = st.verticalAlign
  if (va === 'top') patch.vAlign = 'top'
  else if (va === 'middle') patch.vAlign = 'middle'
  else if (va === 'bottom') patch.vAlign = 'bottom'

  // 自动换行:Excel/WPS 用 white-space 标记 —— 开了换行的格是 `white-space:normal`,默认全局 `td{white-space:nowrap}`。
  // 不读这个 → wrapText 永远 false,长文本不换行而溢出/裁切,连带水平居中也看不出来。只在显式 normal/pre-wrap 时置 true(nowrap/缺省不动)。
  const ws = st.whiteSpace
  if (ws === 'normal' || ws === 'pre-wrap' || ws === 'pre-line') patch.wrapText = true

  // 边框:逐边 border-* 简写;再看 border 通写兜底
  const borders: Record<string, BorderEdge | undefined> = {}
  const sides = [
    ['top', st.borderTop],
    ['right', st.borderRight],
    ['bottom', st.borderBottom],
    ['left', st.borderLeft],
  ] as const
  let anyBorder = false
  const all = st.border ? parseBorderEdge(st.border) : undefined
  for (const [side, spec] of sides) {
    const edge = parseBorderEdge(spec) ?? all
    if (edge) {
      borders[side] = edge
      anyBorder = true
    }
  }
  if (anyBorder) patch.borders = borders

  return Object.keys(patch).length ? patch : null
}

/** 取单元格的值:优先 Excel 的 x:num(数字)/ x:fmla(公式),否则文本(交 setCellValue 推断) */
function cellValueOf(td: Element): CellValue {
  const num = td.getAttribute('x:num')
  if (num != null && num !== '' && !isNaN(Number(num))) return Number(num)
  const fmla = td.getAttribute('x:fmla')
  if (fmla) return fmla[0] === '=' ? fmla : '=' + fmla
  const text = (td.textContent ?? '').replace(/ /g, ' ').trim()
  return text === '' ? null : text
}

/**
 * 解析剪贴板 HTML → {values, styles, merges, images}。非浏览器环境 / 无 <table> 返 null(调用方回退 TSV)。
 */
export function parseClipboardHtml(html: string): ParsedClipboard | null {
  if (typeof DOMParser === 'undefined' || !html) return null
  let doc: Document
  try {
    doc = new DOMParser().parseFromString(html, 'text/html')
  } catch {
    return null
  }
  const table = doc.querySelector('table')
  if (!table) return null

  // Excel/WPS 复制的 HTML 把格式放在 <style> 块的 CSS 类里(<td class="xl65"> + .xl65{...}),
  // 而 td.style 只含内联 style= → 不解析类规则会丢掉绝大多数格式。这里先收集类规则,落格时合并进 td。
  const { classes: classStyles, tdDefault } = parseClassStyles(doc)

  const values: CellValue[][] = []
  const styles: ParsedClipboard['styles'] = []
  const merges: MergeRange[] = []
  const images: ParsedClipboard['images'] = []
  const colWidths: number[] = []
  const rowHeights: number[] = []
  const occupied = new Set<string>() // 被前面 rowspan/colspan 占掉的格 "r:c"

  // 列宽:<col width=N span=M>(N 是 px;Excel/WPS 同时给 style='width:..pt',两者等价,取 px 属性)
  for (const col of Array.from(table.querySelectorAll('col'))) {
    const w = parseInt(col.getAttribute('width') || '', 10)
    const span = Math.max(1, parseInt(col.getAttribute('span') || '1', 10) || 1)
    for (let i = 0; i < span; i++) colWidths.push(Number.isFinite(w) && w > 0 ? w : 0)
  }

  const trs = Array.from(table.querySelectorAll('tr'))
  let r = 0
  for (const tr of trs) {
    const tds = Array.from(tr.children).filter((el) => el.tagName === 'TD' || el.tagName === 'TH')
    if (!tds.length) continue
    // 行高:<tr height=N>(px);缺则看首格 td height
    const trH = parseInt(tr.getAttribute('height') || (tds[0] as Element).getAttribute('height') || '', 10)
    if (Number.isFinite(trH) && trH > 0) rowHeights[r] = trH
    const rowVals: CellValue[] = values[r] ?? (values[r] = [])
    let c = 0
    for (const td of tds) {
      while (occupied.has(`${r}:${c}`)) {
        rowVals[c] = rowVals[c] ?? null
        c++
      }
      rowVals[c] = cellValueOf(td)
      const rawCss = rawCssOf(td as HTMLElement, classStyles, tdDefault) // 合并 td默认+class+内联;返回原始串供解析私有属性
      const patch = cssToStyleOverride(td as HTMLElement) ?? {}
      const numFmt = parseMsoNumberFormat(rawCss) // CSSOM 会丢 mso-*,从原始串解析数字格式(修日期/货币序列号)
      if (numFmt) patch.numFmt = numFmt
      if (Object.keys(patch).length) styles.push({ row: r, col: c, patch })
      // 图片:① data: 的 <img>;② WPS VML o:gfxdata 内嵌图(file:/// 的 <img> 浏览器读不了,只能靠 ②)
      const src = td.querySelector('img')?.getAttribute('src')
      const dataUrl = src && src.startsWith('data:') ? src : extractVmlImageDataUrl(td)
      if (dataUrl) images.push({ row: r, col: c, dataUrl })

      const rs = Math.max(1, parseInt(td.getAttribute('rowspan') || '1', 10) || 1)
      const cs = Math.max(1, parseInt(td.getAttribute('colspan') || '1', 10) || 1)
      if (rs > 1 || cs > 1) {
        merges.push({ top: r, left: c, bottom: r + rs - 1, right: c + cs - 1 })
        for (let dr = 0; dr < rs; dr++)
          for (let dc = 0; dc < cs; dc++) if (dr || dc) occupied.add(`${r + dr}:${c + dc}`)
      }
      c += cs
    }
    r++
  }
  if (!values.length) return null

  // 补齐参差行(右侧空格填 null),稠密对齐
  const width = values.reduce((w, v) => Math.max(w, v.length), 0)
  for (const v of values) for (let i = 0; i < width; i++) if (v[i] === undefined) v[i] = null

  return { values, styles, merges, images, colWidths, rowHeights }
}
