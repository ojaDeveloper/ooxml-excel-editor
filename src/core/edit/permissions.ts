/**
 * 单元格是否可编辑的判定(纯函数,框架无关)。
 * 优先级:总开关关 → 只读;命中 readOnlyRanges → 只读;cellReadOnly 返 true → 只读;否则可编辑。
 */
import type { MergeRange, SheetModel } from '../model/types'
import { cellKey } from '../model/types'
import type { EditConfig } from './types'

function inRange(row: number, col: number, r: MergeRange): boolean {
  return row >= r.top && row <= r.bottom && col >= r.left && col <= r.right
}

export function resolveEditable(sheet: SheetModel, row: number, col: number, cfg: EditConfig): boolean {
  if (!cfg.editable) return false
  if (cfg.readOnlyRanges) {
    for (const r of cfg.readOnlyRanges) if (inRange(row, col, r)) return false
  }
  if (cfg.cellReadOnly) {
    const cell = sheet.cells.get(cellKey(row, col)) ?? null
    if (cfg.cellReadOnly(cell, { row, col })) return false
  }
  return true
}
