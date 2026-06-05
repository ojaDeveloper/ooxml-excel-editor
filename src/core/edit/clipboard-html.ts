/**
 * 解析从 Excel/WPS 复制的剪贴板 HTML(text/html)→ 值 + 样式 + 合并 + 图片。
 * 与 WPS 粘贴一致地还原字体/颜色/填充/边框/对齐/合并;值优先取 Excel 的 x:num/x:fmla(原始值),
 * 否则取单元格文本(交给 setCellValue 推断)。图片仅 data-uri <img>(区域复制一般拿不到内嵌图,见 README)。
 *
 * 用浏览器原生 DOMParser(框架无关;core 仍可 headless,缺 DOMParser 时返 null 回退 TSV)。
 */
import type { BorderEdge, BorderStyle, CellStyleOverride, Font, MergeRange } from '../model/types'
import type { CellValue } from '../model/data-access'
import { toHex6 } from '../format/color'

export interface ParsedClipboard {
  /** 二维值(原始串/数字/公式,交 setCellValue 推断);稠密对齐 */
  values: CellValue[][]
  /** 逐格样式覆盖(只含解析出的字段) */
  styles: { row: number; col: number; patch: CellStyleOverride }[]
  /** 合并区(rowspan/colspan>1) */
  merges: MergeRange[]
  /** data-uri 图片(row,col,dataUrl) */
  images: { row: number; col: number; dataUrl: string }[]
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
    const px = parseFloat(st.fontSize)
    if (px > 0) font.size = Math.round((px * 72) / 96) // px → pt
  }
  if (Object.keys(font).length) patch.font = font

  const bg = toHex6(st.backgroundColor)
  if (bg) patch.fill = { type: 'solid', fgColor: bg }

  if (st.textAlign === 'left' || st.textAlign === 'center' || st.textAlign === 'right') patch.hAlign = st.textAlign
  const va = st.verticalAlign
  if (va === 'top') patch.vAlign = 'top'
  else if (va === 'middle') patch.vAlign = 'middle'
  else if (va === 'bottom') patch.vAlign = 'bottom'

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

  const values: CellValue[][] = []
  const styles: ParsedClipboard['styles'] = []
  const merges: MergeRange[] = []
  const images: ParsedClipboard['images'] = []
  const occupied = new Set<string>() // 被前面 rowspan/colspan 占掉的格 "r:c"

  const trs = Array.from(table.querySelectorAll('tr'))
  let r = 0
  for (const tr of trs) {
    const tds = Array.from(tr.children).filter((el) => el.tagName === 'TD' || el.tagName === 'TH')
    if (!tds.length) continue
    const rowVals: CellValue[] = values[r] ?? (values[r] = [])
    let c = 0
    for (const td of tds) {
      while (occupied.has(`${r}:${c}`)) {
        rowVals[c] = rowVals[c] ?? null
        c++
      }
      rowVals[c] = cellValueOf(td)
      const patch = cssToStyleOverride(td as HTMLElement)
      if (patch) styles.push({ row: r, col: c, patch })
      const img = td.querySelector('img')
      const src = img?.getAttribute('src')
      if (src && src.startsWith('data:')) images.push({ row: r, col: c, dataUrl: src })

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

  return { values, styles, merges, images }
}
