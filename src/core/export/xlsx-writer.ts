/**
 * XLSX 写出(框架无关)—— WorkbookModel → ExcelJS.Workbook → Blob(从编辑后模型重建,所见即所得)。
 * `import('exceljs')` 懒加载(可选 peer,不进 core 产物)。是 exceljs-adapter 读映射的逆。
 *
 * 保真边界(README 同步):重建覆盖 值/公式/样式(字体/填充/边框/对齐/数字格式)/合并/行高列宽/
 * 冻结/图片;**丢失** VBA、工作表保护、复杂 DrawingML(图表/形状按位图另算)、条件格式细节。
 * 接口留口 `fidelity`:日后可加"重载原 ArrayBuffer 叠加"高保真模式,API 不破。
 */
import type { CellModel, CellStyle, ImageAnchor, SheetModel, WorkbookModel } from '../model/types'
import { GridMetrics } from '../layout/grid-metrics'
import { anchorRect } from '../overlay/anchor'
import { emuToPx, DEFAULT_MDW, PX_PER_POINT } from '../layout/units'

export interface XlsxExportOptions {
  /** 保真模式:'rebuild'(默认,从模型重建)。预留 'overlay'(重载原件叠加)日后实现。 */
  fidelity?: 'rebuild'
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
      return { formula: f[0] === '=' ? f.slice(1) : f, result: cell.raw ?? undefined }
    }
    case 'hyperlink':
      return { text: String(cell.raw ?? ''), hyperlink: cell.hyperlink ?? '' }
    case 'richtext':
      return cell.rich ? { richText: cell.rich.map((r) => ({ text: r.text })) } : (cell.raw ?? '')
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
  // 单元格:值 + 样式(空格也套样式,E5 空格上色保真)
  for (const cell of sheet.cells.values()) {
    const ec = ws.getCell(cell.row + 1, cell.col + 1)
    if (cell.type !== 'empty') ec.value = cellValue(cell)
    const st = sheet.styles[cell.styleId]
    if (st) applyStyle(ec, st)
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

/** WorkbookModel → .xlsx Blob(懒加载 exceljs)。 */
export async function workbookToXlsxBlob(workbook: WorkbookModel, _opts: XlsxExportOptions = {}): Promise<Blob> {
  const mod = await import('exceljs')
  const ExcelJS = ((mod as { default?: unknown }).default ?? mod) as { Workbook: new () => any }
  const wb = new ExcelJS.Workbook()
  wb.properties.date1904 = workbook.date1904
  for (const sheet of workbook.sheets) {
    const ws = wb.addWorksheet(sheet.name || `Sheet${sheet.index + 1}`, {
      state: sheet.state === 'visible' ? 'visible' : sheet.state,
    })
    writeSheet(ws, sheet, wb)
  }
  const buf = await wb.xlsx.writeBuffer()
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}
