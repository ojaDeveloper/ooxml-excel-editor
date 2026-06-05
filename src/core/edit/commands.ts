/**
 * 编辑命令(框架无关)。每个命令 apply 时**捕获逆向载荷**,逆向统一为 restore-cells
 * (精确还原一组格的底层 CellModel),于是 undo/redo 跨 值/区域/清空(及后续样式/图片)同一套栈。
 */
import type { CellModel, SheetModel } from '../model/types'
import { cellKey } from '../model/types'
import type { CellValue } from '../model/data-access'
import { cloneCell } from '../model/snapshot'
import { setCellValue, restoreCell } from '../model/mutations'

export type CellPos = { row: number; col: number }

export type EditCommand =
  | { kind: 'set-value'; row: number; col: number; value: CellValue }
  | { kind: 'set-cells'; cells: { row: number; col: number; value: CellValue }[] }
  | { kind: 'restore-cells'; cells: { row: number; col: number; cell: CellModel | null }[] }

export interface ApplyResult {
  /** 该命令的逆命令(undo 时执行,自身又产出 redo 逆命令) */
  inverse: EditCommand
  /** 受影响的格位置(用于发前后快照事件) */
  affected: CellPos[]
}

/** 命令影响到的格位置 */
export function affectedOf(cmd: EditCommand): CellPos[] {
  switch (cmd.kind) {
    case 'set-value':
      return [{ row: cmd.row, col: cmd.col }]
    case 'set-cells':
      return cmd.cells.map((c) => ({ row: c.row, col: c.col }))
    case 'restore-cells':
      return cmd.cells.map((c) => ({ row: c.row, col: c.col }))
  }
}

/** 捕获一组格的当前底层状态(克隆,供逆向精确还原) */
function capture(sheet: SheetModel, positions: CellPos[]): { row: number; col: number; cell: CellModel | null }[] {
  return positions.map((p) => {
    const live = sheet.cells.get(cellKey(p.row, p.col))
    return { row: p.row, col: p.col, cell: live ? cloneCell(live) : null }
  })
}

/** 应用命令并返回逆命令 + 受影响格。 */
export function applyCommand(sheet: SheetModel, cmd: EditCommand): ApplyResult {
  const affected = affectedOf(cmd)
  const prior = capture(sheet, affected)
  switch (cmd.kind) {
    case 'set-value':
      setCellValue(sheet, cmd.row, cmd.col, cmd.value)
      break
    case 'set-cells':
      for (const { row, col, value } of cmd.cells) setCellValue(sheet, row, col, value)
      break
    case 'restore-cells':
      for (const { row, col, cell } of cmd.cells) restoreCell(sheet, row, col, cell)
      break
  }
  return { inverse: { kind: 'restore-cells', cells: prior }, affected }
}
