/**
 * 自动行高: 当某行包含"自动换行"或多行(\n)文字、且按内容换行后超出文件存储的行高时，
 * 把该行撑高到刚好放下内容，避免裁切。
 *
 * 设计取舍(见 EXCEL还原难点.md 第六节):
 * - **只扩不缩**: 最终高度 = max(文件存的高度, 按内容算的高度)，不压用户设定。
 * - **只按"多出来的行"加高**: 第一行锚定在默认/存储高度(那是对的)，第 2 行起才加，
 *   否则会把每个单行单元格都撑高(行盒系数比 Excel 略大)。
 * - **排除合并单元格**: 与 Excel 行为一致(Excel 本就不对合并单元格做自动行高)。
 * - **缺 canvas 环境(node)直接跳过**; WeakSet 防同一 sheet 重复测量。
 *
 * 注意: 这只能挡住"裁切"(难点第 1、2 层)，行高与 Excel 仍可能差几像素(第 3 层不可消除)。
 */
import type { SheetModel, WorkbookModel, CellModel, CellStyle } from '../model/types'
import { cellKey } from '../model/types'
import { formatValue } from '../format/number-format'
import { fontToCss, wrapLines, LINE_HEIGHT_FACTOR, CELL_PADDING } from '../render/text'
import { PX_PER_POINT } from './units'

const fittedSheets = new WeakSet<SheetModel>()

export function autoFitRowHeights(sheet: SheetModel, workbook: WorkbookModel, ctx?: CanvasRenderingContext2D): void {
  if (fittedSheets.has(sheet)) return
  const measure = ctx ?? createMeasureCtx()
  if (!measure) return // node / 无 DOM 环境，跳过
  fittedSheets.add(sheet)

  measure.save()

  // 合并单元格成员(含锚点)全部排除
  const mergedCells = new Set<string>()
  for (const m of sheet.merges) {
    for (let r = m.top; r <= m.bottom; r++) {
      for (let c = m.left; c <= m.right; c++) mergedCells.add(cellKey(r, c))
    }
  }

  const colWidthPx = (c: number): number => {
    const info = sheet.columns.get(c)
    return info?.hidden ? 0 : info?.width ?? sheet.defaultColWidth
  }

  // 每行"额外行带来的最大附加高度(px)"
  const rowExtra = new Map<number, number>()
  for (const [key, cell] of sheet.cells) {
    if (cell.type === 'empty') continue
    if (mergedCells.has(key)) continue
    const style = sheet.styles[cell.styleId]
    const extra = extraHeightOf(measure, cell, style, colWidthPx(cell.col), workbook)
    if (extra > (rowExtra.get(cell.row) ?? 0)) rowExtra.set(cell.row, extra)
  }

  measure.restore()

  // 应用: 自然高度 = 默认行高(第一行) + 额外换行；与存储高度取 max(只扩不缩)
  for (const [r, extra] of rowExtra) {
    if (extra <= 0) continue
    const cur = sheet.rows.get(r)
    if (cur?.hidden) continue
    const base = cur?.height ?? sheet.defaultRowHeight
    const natural = Math.ceil(sheet.defaultRowHeight + extra + CELL_PADDING)
    if (natural > base) sheet.rows.set(r, { height: natural, hidden: false })
  }
}

/** 该单元格因换行多出的高度(px @ zoom=1)。单行返回 0(不撑高)。 */
function extraHeightOf(
  ctx: CanvasRenderingContext2D,
  cell: CellModel,
  style: CellStyle,
  colW: number,
  wb: WorkbookModel,
): number {
  // 快速跳过: 非换行单元格只有"含 \n 的字符串/富文本"才可能多行;
  // 数字/日期/布尔不可能多行 → 连 formatValue 都不用调(大表性能关键)
  if (!style.wrapText && cell.type !== 'string' && cell.type !== 'richtext') return 0

  const text =
    cell.type === 'richtext' && cell.rich
      ? cell.rich.map((r) => r.text).join('')
      : formatValue(cell.raw, style.numFmt, wb.date1904).text
  if (!text) return 0

  let lineCount: number
  if (style.wrapText) {
    const fontCss = fontToCss(style.font, 1)
    const availW = colW - CELL_PADDING * 2
    lineCount = wrapLines(ctx, text, fontCss, availW).length
  } else {
    if (!text.includes('\n')) return 0
    lineCount = text.split('\n').length
  }
  if (lineCount < 2) return 0

  const lineH = style.font.size * PX_PER_POINT * LINE_HEIGHT_FACTOR
  return (lineCount - 1) * lineH
}

function createMeasureCtx(): CanvasRenderingContext2D | null {
  if (typeof document === 'undefined') return null
  const canvas = document.createElement('canvas')
  return canvas.getContext('2d')
}
