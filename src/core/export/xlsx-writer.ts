/**
 * XLSX 写出(框架无关)—— WorkbookModel → ExcelJS.Workbook → Blob(从编辑后模型重建,所见即所得)。
 * `import('exceljs')` 懒加载(可选 peer,不进 core 产物)。是 exceljs-adapter 读映射的逆。
 *
 * 保真边界(README 同步):重建覆盖 值/公式/样式(字体/填充/边框/对齐/数字格式)/合并/行高列宽/
 * 冻结/图片;**丢失** VBA、工作表保护、复杂 DrawingML(图表/形状按位图另算)、条件格式细节。
 * 接口留口 `fidelity`:日后可加"重载原 ArrayBuffer 叠加"高保真模式,API 不破。
 */
import type { CellModel, CellStyle, ConditionalRule, ImageAnchor, MergeRange, SheetModel, WorkbookModel } from '../model/types'
import { cellKey } from '../model/types'
import { GridMetrics } from '../layout/grid-metrics'
import { anchorRect } from '../overlay/anchor'
import { emuToPx, DEFAULT_MDW, PX_PER_POINT } from '../layout/units'
import { injectCellImagesIntoZip } from './wps-cellimages'
import { injectPivotTablesIntoZip, restoreOriginalPivotPartsIntoZip } from './pivot-tables'
import type { ExportProgressFn } from '../progress'
import { checkAborted } from './abort'

export interface XlsxExportOptions {
  /**
   * 保真模式:
   * - `'rebuild'`(默认):从编辑后模型完整重建(干净,但丢原件里我们不建模的部分:条件格式/数据验证/VBA…)。
   * - `'overlay'`:重载原始 .xlsx,只把编辑后的 值/样式/合并/行高列宽/冻结 叠加上去,**保留** ExcelJS
   *   能往返的其余部分(条件格式/数据验证/打印设置/定义名/图表等)。需 `sourceBuffer`(壳自动注入);
   *   缺原件时自动回退 rebuild。注:overlay 不反映 结构增删行列 / 图片 编辑(那类用 rebuild)。
   *
   * WPS 单元格内嵌图(DISPIMG):两种模式都会在 ExcelJS 写出后**于 zip 层回注** `xl/cellimages.xml`
   * + rels + media + Content_Types/workbook-rels(见 wps-cellimages.ts),从模型重建 → 原有的 + App 内
   * 新转的内嵌图导出后都在(blob-only 无字节的图除外)。导出 → 用 WPS 打开,内嵌图正常显示。
   */
  fidelity?: 'rebuild' | 'overlay'
  /** 原始 .xlsx 字节(overlay 模式用;由 exporter 从 host 注入,用方一般不直接传) */
  sourceBuffer?: ArrayBuffer
  /**
   * 透视表零件回注开关(默认 false,经 viewer 导出时随 `EditConfig.pivotTable` 自动注入)。
   * 开启时:① App 内创建的透视表重建成真实 OOXML 零件(pivot-tables.ts);② overlay 模式
   * 从 `sourceBuffer` 原样搬运原文件的透视表零件(ExcelJS 不建模 pivot,不搬运就丢)。
   */
  pivotTables?: boolean
  /** 长任务进度回调(zip 写出前/后 emit `{stage:'zip'}`;exceljs writeBuffer 黑盒) */
  onProgress?: ExportProgressFn
  /** 取消信号(zip 阶段前后检查) */
  signal?: AbortSignal
}

/** css 颜色 → ExcelJS ARGB('FFRRGGBB');无法识别返 undefined。 */
function cssToArgb(c?: string): string | undefined {
  if (!c) return undefined
  let m = /^#([0-9a-f]{6})$/i.exec(c)
  if (m) return ('FF' + m[1]).toUpperCase()
  m = /^#([0-9a-f]{8})$/i.exec(c) // #RRGGBBAA → AARRGGBB
  if (m) return (m[1].slice(6) + m[1].slice(0, 6)).toUpperCase()
  const rgba = /rgba?\(([^)]+)\)/i.exec(c)
  if (rgba) {
    const p = rgba[1].split(',').map((s) => s.trim())
    const hex = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0')
    const a = hex((p[3] != null ? +p[3] : 1) * 255)
    return (a + hex(+p[0]) + hex(+p[1]) + hex(+p[2])).toUpperCase()
  }
  return undefined
}

/** 把模型样式套到 ExcelJS cell(只设非默认部分,保文件精简)。 */
/** 我们的 Partial<Font> → ExcelJS 富文本 run 字体(只写出现的字段);颜色 css→argb。 */
function toExcelRichFont(f: Partial<CellStyle['font']>): Record<string, unknown> {
  const color = f.color ? cssToArgb(f.color) : undefined
  const out: Record<string, unknown> = {}
  if (f.name) out.name = f.name
  if (f.size) out.size = f.size
  if (f.bold) out.bold = true
  if (f.italic) out.italic = true
  if (f.underline) out.underline = true
  if (f.strike) out.strike = true
  if (color) out.color = { argb: color }
  return out
}

function applyStyle(ec: { font?: unknown; fill?: unknown; border?: unknown; alignment?: unknown; numFmt?: string }, st: CellStyle): void {
  const f = st.font
  const color = cssToArgb(f.color)
  ec.font = {
    name: f.name || undefined,
    size: f.size || undefined,
    bold: f.bold || undefined,
    italic: f.italic || undefined,
    underline: f.underline || undefined,
    strike: f.strike || undefined,
    color: color ? { argb: color } : undefined,
  }
  if (st.fill && st.fill.type !== 'none') {
    const fg = cssToArgb(st.fill.fgColor)
    const bg = cssToArgb(st.fill.bgColor)
    ec.fill = {
      type: 'pattern',
      pattern: st.fill.type === 'solid' ? 'solid' : st.fill.pattern || 'solid',
      fgColor: fg ? { argb: fg } : undefined,
      bgColor: bg ? { argb: bg } : undefined,
    }
  }
  const b = st.borders
  if (b && (b.top || b.bottom || b.left || b.right || b.diagonal)) {
    const edge = (e?: { style: string; color: string }) =>
      e && e.style !== 'none' ? { style: e.style, color: cssToArgb(e.color) ? { argb: cssToArgb(e.color) } : undefined } : undefined
    ec.border = {
      top: edge(b.top),
      bottom: edge(b.bottom),
      left: edge(b.left),
      right: edge(b.right),
      diagonal: b.diagonal ? { up: !!b.diagonalUp, down: !!b.diagonalDown, ...edge(b.diagonal) } : undefined,
    }
  }
  const align: Record<string, unknown> = {}
  if (st.hAlign && st.hAlign !== 'general') align.horizontal = st.hAlign
  if (st.vAlign && st.vAlign !== 'bottom') align.vertical = st.vAlign
  if (st.wrapText) align.wrapText = true
  if (st.textRotation) align.textRotation = st.textRotation
  if (st.indent) align.indent = st.indent
  if (Object.keys(align).length) ec.alignment = align
  if (st.numFmt && st.numFmt !== 'General') ec.numFmt = st.numFmt
}

// ===== 条件格式导出(1.9.0) =====
// 策略(rebuild + overlay 统一):清空 ExcelJS 现有 CF(overlay 重载原件带进来的),按模型重建。
//  - parsed 且未编辑(origin!=='user' && !dirty)→ 原样回写解析时存下的原始 ExcelJS rule(raw),
//    含 cfvo 阈值等我们不全建模的字段 → 未动的原规则零退化。
//  - user 新建 / 编辑过(dirty)→ 按模型 buildExcelCfRule 重建。

function colLetters(col: number): string {
  let s = ''
  let n = col
  do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1 } while (n >= 0)
  return s
}
function rangeRef(r: MergeRange): string {
  return `${colLetters(r.left)}${r.top + 1}:${colLetters(r.right)}${r.bottom + 1}`
}
function rangesToRef(ranges: MergeRange[]): string {
  return (ranges || []).map(rangeRef).join(' ')
}

/** 我们的条件格式样式 → ExcelJS 条件格式 dxf 样式(font/fill);只写出现的字段。 */
function cfStyleToExcel(style?: ConditionalRule['style']): Record<string, unknown> | undefined {
  if (!style) return undefined
  const out: Record<string, unknown> = {}
  const f = style.font
  if (f) {
    const font: Record<string, unknown> = {}
    if (f.bold) font.bold = true
    if (f.italic) font.italic = true
    if (f.underline) font.underline = true
    if (f.strike) font.strike = true
    const color = cssToArgb(f.color)
    if (color) font.color = { argb: color }
    if (Object.keys(font).length) out.font = font
  }
  if (style.fill && style.fill.type !== 'none') {
    const fg = cssToArgb(style.fill.fgColor)
    if (fg) out.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fg } }
  }
  return Object.keys(out).length ? out : undefined
}

/** 模型规则 → ExcelJS addConditionalFormatting 的 rule(用户新建/编辑过的规则用)。无法构造返 null。 */
function buildExcelCfRule(rule: ConditionalRule): Record<string, unknown> | null {
  const priority = rule.priority || 1
  const style = cfStyleToExcel(rule.style)
  switch (rule.type) {
    case 'cellIs':
      return { type: 'cellIs', operator: rule.operator || 'greaterThan', formulae: rule.formulae || [], priority, style }
    case 'expression':
      return { type: 'expression', formulae: rule.formulae || [], priority, style }
    case 'colorScale': {
      const cs = rule.colorScale
      if (!cs) return null
      const cfvo = cs.mid ? [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }]
      const cols = (cs.mid ? [cs.min, cs.mid, cs.max] : [cs.min, cs.max]).map((c) => ({ argb: cssToArgb(c) || 'FF000000' }))
      return { type: 'colorScale', cfvo, color: cols, priority }
    }
    case 'dataBar': {
      const db = rule.dataBar
      if (!db) return null
      return { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: cssToArgb(db.color) || 'FF638EC6' }, priority }
    }
    case 'iconSet': {
      const ic = rule.iconSet
      if (!ic) return null
      return { type: 'iconSet', iconSet: ic.name, reverse: !!ic.reverse, cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }], priority }
    }
    case 'top10': {
      const t = rule.top10 || { rank: 10, percent: false, bottom: false }
      return { type: 'top10', rank: t.rank, percent: t.percent, bottom: t.bottom, priority, style }
    }
    default:
      return null
  }
}

/** 把模型的条件格式写到 ExcelJS 工作表(rebuild + overlay 共用)。 */
function writeConditionalFormatting(ws: any, sheet: SheetModel): void {
  const rules = sheet.conditional
  // overlay 重载原件后 ws 上已有原 CF;无论有没有规则都先清空,再按模型统一重建(含"用户删光"的情况)
  if (Array.isArray(ws.conditionalFormattings)) ws.conditionalFormattings = []
  if (!rules || !rules.length) return
  for (const rule of rules) {
    if (rule.type === 'unsupported') continue
    const ref = rangesToRef(rule.ranges)
    if (!ref) continue
    const useRaw = rule.origin !== 'user' && !rule.dirty && rule.raw
    const built = useRaw ? (rule.raw as Record<string, unknown>) : buildExcelCfRule(rule)
    if (!built) continue
    try {
      ws.addConditionalFormatting({ ref, rules: [built] })
    } catch {
      /* 单条规则构造失败不致命,跳过 */
    }
  }
}

/** 模型 cell → ExcelJS value(按类型)。 */
function cellValue(cell: CellModel): unknown {
  switch (cell.type) {
    case 'number':
    case 'boolean':
    case 'date':
    case 'string':
      return cell.raw
    case 'error':
      return { error: cell.raw }
    case 'formula': {
      const f = cell.formula ?? ''
      const formula = f[0] === '=' ? f.slice(1) : f
      // WPS DISPIMG 内嵌图格:缓存值 <v> 写成 =DISPIMG("id",1)(对齐真·WPS;图能加载时不显示此值,加载失败时作回退)
      if (cell.dispImgId) return { formula, result: `=DISPIMG("${cell.dispImgId}",1)` }
      return { formula, result: cell.raw ?? undefined }
    }
    case 'hyperlink':
      return { text: String(cell.raw ?? ''), hyperlink: cell.hyperlink ?? '' }
    case 'richtext':
      // 带回每段字体(颜色/粗斜/下划线/删除线/字号/字体名)→ 富文本往返不丢格式
      return cell.rich ? { richText: cell.rich.map((r) => (r.font ? { text: r.text, font: toExcelRichFont(r.font) } : { text: r.text })) } : (cell.raw ?? '')
    default:
      return cell.raw ?? null // empty
  }
}

function extOfMime(mime?: string): 'png' | 'jpeg' | 'gif' {
  if (mime?.includes('jpeg') || mime?.includes('jpg')) return 'jpeg'
  if (mime?.includes('gif')) return 'gif'
  return 'png'
}

/** 落图片到 ExcelJS(bytes → buffer / data url → base64;blob: url 无法同步读 → 跳过)。 */
function addImageTo(wb: { addImage(o: unknown): number }, ws: { addImage(id: number, a: unknown): void }, anchor: ImageAnchor, metrics: GridMetrics): void {
  let imgOpts: unknown
  if (anchor.bytes) imgOpts = { buffer: anchor.bytes, extension: extOfMime(anchor.mime) }
  else if (anchor.src?.startsWith('data:')) {
    const comma = anchor.src.indexOf(',')
    const meta = anchor.src.slice(5, comma)
    imgOpts = { base64: anchor.src.slice(comma + 1), extension: extOfMime(meta) }
  } else return // blob: / http url 同步读不了,跳过
  const id = wb.addImage(imgOpts)
  // 单元格 + EMU 偏移 → ExcelJS 分数列/行(整数部=格,小数部=该格内偏移比例;ExcelJS 据列宽换回 EMU)
  const fracCol = (col: number, offEmu: number) => fractional((c) => metrics.colWidth(c), col, emuToPx(offEmu))
  const fracRow = (row: number, offEmu: number) => fractional((r) => metrics.rowHeight(r), row, emuToPx(offEmu))
  const tl = { col: fracCol(anchor.from.col, anchor.from.colOffEmu), row: fracRow(anchor.from.row, anchor.from.rowOffEmu) }
  const editAs =
    anchor.editAs === 'oneCell' || anchor.editAs === 'absolute' || anchor.editAs === 'twoCell'
      ? anchor.editAs
      : anchor.to
        ? 'twoCell'
        : 'oneCell'
  if (anchor.to) {
    // 双格锚:tl + br(随单元格缩放,保真原始 twoCellAnchor)
    ws.addImage(id, { tl, br: { col: fracCol(anchor.to.col, anchor.to.colOffEmu), row: fracRow(anchor.to.row, anchor.to.rowOffEmu) }, editAs })
  } else {
    // 单格锚:tl + 像素尺寸(origin 归一/oneCellAnchor)
    const rect = anchorRect(metrics, anchor)
    ws.addImage(id, { tl, ext: { width: rect.width, height: rect.height }, editAs })
  }
}

/** 单元格 base + 像素偏移 → ExcelJS 分数索引(走列/行宽累减,跨格进位)。 */
function fractional(sizeAt: (i: number) => number, base: number, offPx: number): number {
  let i = base
  let off = offPx
  while (off > 0) {
    const sz = sizeAt(i)
    if (sz <= 0 || off < sz) break
    off -= sz
    i++
  }
  const sz = sizeAt(i)
  return i + (sz > 0 ? off / sz : 0)
}

/** 重建一个工作表到 ExcelJS。 */
function writeSheet(ws: any, sheet: SheetModel, wb: any): void {
  // 列宽(px → 字符)/隐藏
  for (const [c, info] of sheet.columns) {
    const col = ws.getColumn(c + 1)
    col.width = Math.max(0, (info.width - 5) / DEFAULT_MDW)
    if (info.hidden) col.hidden = true
  }
  // 行高(px → pt)/隐藏
  for (const [r, info] of sheet.rows) {
    const row = ws.getRow(r + 1)
    row.height = info.height / PX_PER_POINT
    if (info.hidden) row.hidden = true
  }
  // 单元格:值 + 样式(空格也套样式,E5 空格上色保真)+ 批注(1.11.0)
  for (const cell of sheet.cells.values()) {
    const ec = ws.getCell(cell.row + 1, cell.col + 1)
    if (cell.type !== 'empty') ec.value = cellValue(cell)
    const st = sheet.styles[cell.styleId]
    if (st) applyStyle(ec, st)
    if (cell.comment) ec.note = cell.comment
  }
  // 合并(1-based;重叠/越界吞掉不致命)
  for (const m of sheet.merges) {
    try {
      ws.mergeCells(m.top + 1, m.left + 1, m.bottom + 1, m.right + 1)
    } catch {
      /* 跳过非法合并 */
    }
  }
  // 冻结窗格
  if (sheet.freeze && (sheet.freeze.frozenRows || sheet.freeze.frozenCols)) {
    ws.views = [{ state: 'frozen', xSplit: sheet.freeze.frozenCols, ySplit: sheet.freeze.frozenRows }]
  }
  // 条件格式(1.9.0:rebuild 也回写,从模型重建;parsed 未编辑的用 raw 原样回写)
  writeConditionalFormatting(ws, sheet)
  // 图片(best-effort,单张失败跳过)
  const metrics = new GridMetrics(sheet, 1)
  for (const anchor of sheet.images) {
    try {
      addImageTo(wb, ws, anchor, metrics)
    } catch {
      /* 跳过单张图 */
    }
  }
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

/** overlay:把编辑后模型的 值/样式/合并/行高列宽/冻结 叠加到已加载的原始工作表(保留其余原件部分)。 */
function applyModelOntoSheet(ws: any, sheet: SheetModel): void {
  const live = new Set(sheet.cells.keys())
  // 清掉原件里有、模型里没有的格(被删/清空)
  ws.eachRow({ includeEmpty: false }, (row: any, rNum: number) => {
    row.eachCell({ includeEmpty: false }, (cell: any, cNum: number) => {
      if (!live.has(cellKey(rNum - 1, cNum - 1))) cell.value = null
    })
  })
  // 套模型格(值 + 样式 + 批注)
  for (const cell of sheet.cells.values()) {
    const ec = ws.getCell(cell.row + 1, cell.col + 1)
    ec.value = cell.type !== 'empty' ? cellValue(cell) : null
    const st = sheet.styles[cell.styleId]
    if (st) applyStyle(ec, st)
    if (cell.comment) ec.note = cell.comment
  }
  // 合并:先拆原有,再按模型合
  for (const m of [...(ws.model?.merges ?? [])]) {
    try {
      ws.unMergeCells(m)
    } catch {
      /* 忽略 */
    }
  }
  for (const m of sheet.merges) {
    try {
      ws.mergeCells(m.top + 1, m.left + 1, m.bottom + 1, m.right + 1)
    } catch {
      /* 忽略 */
    }
  }
  for (const [c, info] of sheet.columns) {
    const col = ws.getColumn(c + 1)
    col.width = Math.max(0, (info.width - 5) / DEFAULT_MDW)
    if (info.hidden) col.hidden = true
  }
  for (const [r, info] of sheet.rows) {
    const row = ws.getRow(r + 1)
    row.height = info.height / PX_PER_POINT
    if (info.hidden) row.hidden = true
  }
  if (sheet.freeze && (sheet.freeze.frozenRows || sheet.freeze.frozenCols)) {
    ws.views = [{ state: 'frozen', xSplit: sheet.freeze.frozenCols, ySplit: sheet.freeze.frozenRows }]
  }
  // 条件格式(1.9.0:overlay 也按模型重建 —— 未编辑的 parsed 规则用 raw 原样回写,不退化;
  // 用户新建/编辑/删除的反映进去)。注:清空 ws 原 CF 再重建,所以即便全删也对。
  writeConditionalFormatting(ws, sheet)
}

/**
 * WorkbookModel → .xlsx 原始字节(`Uint8Array`,懒加载 exceljs)。overlay 模式重载原件叠加编辑,
 * 否则从模型重建。**框架无关、纯 Node 可用**(无 DOM/Blob 依赖)——
 * Node 侧:`fs.writeFileSync(out, await workbookToXlsxBytes(wb, { fidelity: 'overlay', sourceBuffer }))`。
 * 浏览器侧若要触发下载,用 `workbookToXlsxBlob`(它就是本函数 + Blob 包装)。
 */
export async function workbookToXlsxBytes(workbook: WorkbookModel, opts: XlsxExportOptions = {}): Promise<Uint8Array> {
  checkAborted(opts.signal)
  const mod = await import('exceljs')
  const ExcelJS = ((mod as { default?: unknown }).default ?? mod) as { Workbook: new () => any }

  // 高保真 overlay:重载原件 + 叠加编辑(保留 ExcelJS 能往返的原件部分)
  if (opts.fidelity === 'overlay' && opts.sourceBuffer) {
    checkAborted(opts.signal)
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(opts.sourceBuffer as any)
    for (let i = 0; i < workbook.sheets.length; i++) {
      checkAborted(opts.signal)
      const sheet = workbook.sheets[i]
      const ws = wb.worksheets[i] ?? wb.getWorksheet(sheet.name) ?? wb.addWorksheet(sheet.name || `Sheet${i + 1}`)
      applyModelOntoSheet(ws, sheet)
      opts.onProgress?.({ stage: 'write', sheetIndex: i, ratio: (i + 1) / workbook.sheets.length, label: `写入 ${sheet.name}` })
    }
    opts.onProgress?.({ stage: 'zip', ratio: undefined, label: 'zip 压缩 .xlsx…' })
    const buf = await wb.xlsx.writeBuffer()
    checkAborted(opts.signal)
    opts.onProgress?.({ stage: 'zip', ratio: 1 })
    return finalizeBytes(buf, workbook, opts)
  }

  // 默认:从模型完整重建
  const wb = new ExcelJS.Workbook()
  wb.properties.date1904 = workbook.date1904
  for (let i = 0; i < workbook.sheets.length; i++) {
    checkAborted(opts.signal)
    const sheet = workbook.sheets[i]
    const ws = wb.addWorksheet(sheet.name || `Sheet${sheet.index + 1}`, {
      state: sheet.state === 'visible' ? 'visible' : sheet.state,
    })
    writeSheet(ws, sheet, wb)
    opts.onProgress?.({ stage: 'write', sheetIndex: i, ratio: (i + 1) / workbook.sheets.length, label: `写入 ${sheet.name}` })
  }
  opts.onProgress?.({ stage: 'zip', ratio: 0, label: 'zip 压缩…' })
  const buf = await wb.xlsx.writeBuffer()
  checkAborted(opts.signal)
  opts.onProgress?.({ stage: 'zip', ratio: 1 })
  return finalizeBytes(buf, workbook, opts)
}

/** WorkbookModel → .xlsx Blob(浏览器下载用)。= `workbookToXlsxBytes` + Blob 包装,签名/行为不变。 */
export async function workbookToXlsxBlob(workbook: WorkbookModel, opts: XlsxExportOptions = {}): Promise<Blob> {
  const bytes = await workbookToXlsxBytes(workbook, opts)
  return new Blob([bytes as BlobPart], { type: XLSX_MIME })
}

/**
 * ExcelJS 写出后:回注 WPS 单元格内嵌图(DISPIMG)私有件 + 真实 OOXML 透视表零件,返回最终字节。
 * 透视表回注受 `opts.pivotTables` 开关控制(默认关):先搬运原文件零件(overlay 有原件时),
 * 再注 App 内创建的(编号自动避开已搬运的)。都没有时零开销。**纯字节运算,无浏览器依赖。**
 */
function finalizeBytes(buf: ArrayBuffer, workbook: WorkbookModel, opts: XlsxExportOptions): Uint8Array {
  let bytes: Uint8Array = new Uint8Array(buf)
  if (workbook.cellImages && workbook.cellImages.size) {
    bytes = injectCellImagesIntoZip(bytes, workbook)
  }
  if (opts.pivotTables) {
    try {
      if (opts.fidelity === 'overlay' && opts.sourceBuffer) {
        bytes = restoreOriginalPivotPartsIntoZip(bytes, new Uint8Array(opts.sourceBuffer))
      }
      bytes = injectPivotTablesIntoZip(bytes, workbook)
    } catch {
      /* 透视表回注失败不影响主体导出(静态结果仍在单元格里) */
    }
  }
  return bytes
}
