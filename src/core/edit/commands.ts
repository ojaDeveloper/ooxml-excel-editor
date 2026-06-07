/**
 * 编辑命令(框架无关)。每个命令 apply 时**捕获逆向载荷**,逆向统一为 restore-cells
 * (精确还原一组格的底层 CellModel),于是 undo/redo 跨 值/区域/清空(及后续样式/图片)同一套栈。
 */
import type { CellModel, CellStyleOverride, ColumnInfo, ImageAnchor, MergeRange, RowInfo, SheetModel, WorkbookModel } from '../model/types'
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
  addImage,
  removeImage,
  cloneImageAnchor,
} from '../model/mutations'
import type { StructOp } from '../model/structure'

export type CellPos = { row: number; col: number }
export type DimAxis = 'col' | 'row'

export type EditCommand =
  | { kind: 'set-value'; row: number; col: number; value: CellValue }
  | { kind: 'set-cells'; cells: { row: number; col: number; value: CellValue }[] }
  | { kind: 'restore-cells'; cells: { row: number; col: number; cell: CellModel | null }[] }
  | { kind: 'set-dim'; axis: DimAxis; index: number; size: number }
  | { kind: 'restore-dim'; axis: DimAxis; index: number; info: ColumnInfo | RowInfo | null }
  | { kind: 'set-style'; cells: CellPos[]; patch: CellStyleOverride }
  | { kind: 'image-set'; index: number; anchor: ImageAnchor }
  | { kind: 'image-add'; anchor: ImageAnchor; index?: number }
  | { kind: 'image-remove'; index: number }
  | { kind: 'struct-edit'; op: StructOp; at: number; count: number }
  | { kind: 'restore-wb'; snapshot: WorkbookModel }
  | { kind: 'merge-cells'; range: MergeRange }
  | { kind: 'unmerge-cells'; range: MergeRange }
  | { kind: 'restore-merges'; merges: MergeRange[]; cells: { row: number; col: number; cell: CellModel | null }[] }
  // WPS 内嵌图 ⇄ 浮动图互转(第二期):由 EditController.exec 直接处理(需 workbook 级快照逆),不走 applyCommand
  | { kind: 'convert-to-cell'; imageIndex: number; row: number; col: number }
  | { kind: 'convert-to-cells'; targets: { imageIndex: number; row: number; col: number }[] }
  | { kind: 'convert-to-float'; row: number; col: number; size?: { width: number; height: number } }
  | { kind: 'convert-to-floats'; cells: { row: number; col: number; size?: { width: number; height: number } }[] }

/** dim 命令(列宽/行高)— 仅维度族,无格位置 */
export type DimCommand = Extract<EditCommand, { kind: 'set-dim' } | { kind: 'restore-dim' }>
export function isDimCommand(cmd: EditCommand): cmd is DimCommand {
  return cmd.kind === 'set-dim' || cmd.kind === 'restore-dim'
}

/** image 命令(浮动/嵌入图片增删移改)— 无格位置,发 image-change */
export type ImageCommand = Extract<EditCommand, { kind: 'image-set' } | { kind: 'image-add' } | { kind: 'image-remove' }>
export function isImageCommand(cmd: EditCommand): cmd is ImageCommand {
  return cmd.kind === 'image-set' || cmd.kind === 'image-add' || cmd.kind === 'image-remove'
}

/** struct 命令(增删行列)— 无格位置,发 struct-change,整簿快照逆(支持跨表公式重写撤销) */
export type StructCommand = Extract<EditCommand, { kind: 'struct-edit' } | { kind: 'restore-wb' }>
export function isStructCommand(cmd: EditCommand): cmd is StructCommand {
  return cmd.kind === 'struct-edit' || cmd.kind === 'restore-wb'
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
    case 'merge-cells':
      return coveredOf(cmd.range) // 合并清空的被覆盖格 → 发 cell-change
    case 'restore-merges':
      return cmd.cells.map((c) => ({ row: c.row, col: c.col }))
    case 'set-dim':
    case 'restore-dim':
    case 'image-set':
    case 'image-add':
    case 'image-remove':
    case 'struct-edit':
    case 'restore-wb':
    case 'unmerge-cells':
    case 'convert-to-cell':
    case 'convert-to-cells':
    case 'convert-to-float':
    case 'convert-to-floats':
      return []
  }
}

/** 区域内"被覆盖格"(除左上锚点);合并时这些格的值被清空。 */
function coveredOf(range: MergeRange): CellPos[] {
  const out: CellPos[] = []
  for (let r = range.top; r <= range.bottom; r++)
    for (let c = range.left; c <= range.right; c++) if (!(r === range.top && c === range.left)) out.push({ row: r, col: c })
  return out
}
function mergesIntersect(a: MergeRange, b: MergeRange): boolean {
  return !(a.bottom < b.top || a.top > b.bottom || a.right < b.left || a.left > b.right)
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
  // 图片族:逆=对偶命令(set↔set / add↔remove)
  if (cmd.kind === 'image-set') {
    const prior = cloneImageAnchor(sheet.images[cmd.index])
    sheet.images[cmd.index] = cloneImageAnchor(cmd.anchor)
    return { inverse: { kind: 'image-set', index: cmd.index, anchor: prior }, affected }
  }
  if (cmd.kind === 'image-add') {
    const at = addImage(sheet, cloneImageAnchor(cmd.anchor), cmd.index)
    return { inverse: { kind: 'image-remove', index: at }, affected }
  }
  if (cmd.kind === 'image-remove') {
    const prior = cloneImageAnchor(sheet.images[cmd.index])
    removeImage(sheet, cmd.index)
    return { inverse: { kind: 'image-add', anchor: prior, index: cmd.index }, affected }
  }
  // 合并族:逆=restore-merges(整张 merges 数组 + 被清空格的前态),覆盖 合并/拆分/还原 三向
  if (cmd.kind === 'merge-cells') {
    const priorMerges = sheet.merges.map((m) => ({ ...m }))
    const covered = coveredOf(cmd.range)
    const priorCells = capture(sheet, covered)
    sheet.merges = sheet.merges.filter((m) => !mergesIntersect(m, cmd.range)) // 吸收相交的旧合并
    for (const p of covered) sheet.cells.delete(cellKey(p.row, p.col)) // 清空被覆盖格(只留锚点)
    sheet.merges.push({ ...cmd.range })
    return { inverse: { kind: 'restore-merges', merges: priorMerges, cells: priorCells }, affected }
  }
  if (cmd.kind === 'unmerge-cells') {
    const priorMerges = sheet.merges.map((m) => ({ ...m }))
    sheet.merges = sheet.merges.filter((m) => !mergesIntersect(m, cmd.range))
    return { inverse: { kind: 'restore-merges', merges: priorMerges, cells: [] }, affected }
  }
  if (cmd.kind === 'restore-merges') {
    const curMerges = sheet.merges.map((m) => ({ ...m }))
    const curCells = capture(sheet, cmd.cells.map((c) => ({ row: c.row, col: c.col })))
    sheet.merges = cmd.merges.map((m) => ({ ...m }))
    for (const { row, col, cell } of cmd.cells) restoreCell(sheet, row, col, cell)
    return { inverse: { kind: 'restore-merges', merges: curMerges, cells: curCells }, affected }
  }
  // 结构族(增删行列)由 EditController.exec 直接处理(需 workbook 级快照 + 跨表公式重写),不走这里
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
