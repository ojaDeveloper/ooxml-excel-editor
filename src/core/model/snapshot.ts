/**
 * CellSnapshot —— 单元格的"前后完整快照"(编辑事件载荷 + 查询 API 的统一底层结构)。
 * 复用读层单一真相源:getCell + getCellStyle + cellDisplayText,不另写取值/格式逻辑。
 *
 * 字段语义:
 * - cell:    底层 CellModel 的克隆(含 type/raw/formula/rich/hyperlink/comment/styleId),整个结构都给出。
 * - style:   解析后的 CellStyle(由 styleId 解引用)。
 * - raw:     原始存值(cell.raw);公式格里 = 缓存结果。
 * - computed:有效计算值(当前 = raw;E4 接公式引擎后,引擎结果也写回 raw,故仍 = raw)。
 * - text:    显示文本(cellDisplayText,与渲染所见一致)。
 */
import type { CellModel, CellStyle, SheetModel } from './types'
import { cellKey } from './types'
import type { CellValue } from './data-access'
import { cellDisplayText, getCellStyle } from './data-access'

export interface CellSnapshot {
  row: number
  col: number
  cell: CellModel | null
  style: CellStyle | undefined
  raw: CellValue
  computed: CellValue
  text: string
}

/** 深克隆 CellModel(raw 的 Date / rich 数组都拷,避免前后快照共享引用) */
export function cloneCell(c: CellModel): CellModel {
  return {
    ...c,
    raw: c.raw instanceof Date ? new Date(c.raw.getTime()) : c.raw,
    rich: c.rich ? c.rich.map((r) => ({ ...r, font: r.font ? { ...r.font } : undefined })) : undefined,
  }
}

export function buildCellSnapshot(sheet: SheetModel, row: number, col: number, date1904: boolean): CellSnapshot {
  const live = sheet.cells.get(cellKey(row, col))
  const style = getCellStyle(sheet, row, col)
  const raw = live && live.type !== 'empty' ? live.raw : null
  return {
    row,
    col,
    cell: live ? cloneCell(live) : null,
    style,
    raw,
    computed: raw, // 当前 = raw;E4 引擎结果也写回 raw,语义不变
    text: cellDisplayText(live, style, date1904),
  }
}
