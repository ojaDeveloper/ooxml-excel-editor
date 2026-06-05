/**
 * 编辑命令(框架无关)。每个命令 apply 时**捕获逆向载荷**,逆向统一为 restore-cells
 * (精确还原一组格的底层 CellModel),于是 undo/redo 跨 值/区域/清空(及后续样式/图片)同一套栈。
 */
import type { CellModel, CellStyleOverride, ColumnInfo, RowInfo, SheetModel } from '../model/types'
import { cellKey } from '../model/types'
import type { CellValue } from '../model/data-access'
import { cloneCell } from '../model/snapshot'
import {
  setCellValue,
  restoreCell,
  setColumnWidth,
  setRowHeight,
  restoreDimension,
  applyStyleOverride,
} from '../model/mutations'

export type CellPos = { row: number; col: number }
export type DimAxis = 'col' | 'row'

export type EditCommand =
  | { kind: 'set-value'; row: number; col: number; value: CellValue }
  | { kind: 'set-cells'; cells: { row: number; col: number; value: CellValue }[] }
  | { kind: 'restore-cells'; cells: { row: number; col: number; cell: CellModel | null }[] }
  | { kind: 'set-dim'; axis: DimAxis; index: number; size: number }
  | { kind: 'restore-dim'; axis: DimAxis; index: number; info: ColumnInfo | RowInfo | null }
  | { kind: 'set-style'; cells: CellPos[]; patch: CellStyleOverride }

/** dim 命令(列宽/行高)— 仅维度族,无格位置 */
export type DimCommand = Extract<EditCommand, { kind: 'set-dim' } | { kind: 'restore-dim' }>
export function isDimCommand(cmd: EditCommand): cmd is DimCommand {
  return cmd.kind === 'set-dim' || cmd.kind === 'restore-dim'
}

export interface ApplyResult {
  /** 该命令的逆命令(undo 时执行,自身又产出 redo 逆命令) */
  inverse: EditCommand
  /** 受影响的格位置(用于发前后快照事件) */
  affected: CellPos[]
}

/** 命令影响到的格位置(dim 命令无格位置 → 返 []) */
export function affectedOf(cmd: EditCommand): CellPos[] {
  switch (cmd.kind) {
    case 'set-value':
      return [{ row: cmd.row, col: cmd.col }]
    case 'set-cells':
      return cmd.cells.map((c) => ({ row: c.row, col: c.col }))
    case 'restore-cells':
      return cmd.cells.map((c) => ({ row: c.row, col: c.col }))
    case 'set-style':
      return cmd.cells.map((c) => ({ row: c.row, col: c.col }))
    case 'set-dim':
    case 'restore-dim':
      return []
  }
}

/** 捕获一个列/行当前维度信息(克隆,供逆向还原;无项 → null) */
function captureDim(sheet: SheetModel, axis: DimAxis, index: number): ColumnInfo | RowInfo | null {
  const info = (axis === 'col' ? sheet.columns : sheet.rows).get(index)
  return info ? { ...info } : null
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
  // 维度族:逆=restore-dim(捕获前态),与 cell 族的 restore-cells 同构
  if (cmd.kind === 'set-dim' || cmd.kind === 'restore-dim') {
    const priorInfo = captureDim(sheet, cmd.axis, cmd.index)
    if (cmd.kind === 'set-dim') {
      if (cmd.axis === 'col') setColumnWidth(sheet, cmd.index, cmd.size)
      else setRowHeight(sheet, cmd.index, cmd.size)
    } else {
      restoreDimension(sheet, cmd.axis, cmd.index, cmd.info)
    }
    return { inverse: { kind: 'restore-dim', axis: cmd.axis, index: cmd.index, info: priorInfo }, affected }
  }
  // 单元格族:逆=restore-cells(捕获前态;set-style 也走它 → 空格上色的逆=删格)
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
    case 'set-style':
      for (const { row, col } of cmd.cells) applyStyleOverride(sheet, row, col, cmd.patch)
      break
  }
  return { inverse: { kind: 'restore-cells', cells: prior }, affected }
}
