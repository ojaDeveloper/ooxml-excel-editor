/**
 * 行列结构编辑(框架无关)—— 插入/删除 行/列。重键 cells Map、移合并区、移行高列宽、移图片锚点、
 * 调 dimension。撤销走"结构快照"逆命令:命令应用前拍整张结构快照,undo 整体还原(简单且铁正确)。
 *
 * v1 限制:不重写公式引用文本(=A5 删一行后仍是 =A5,缓存值随格移动)。开 recalc 时引擎在结构变更后
 * 重建(refreshEngine),后续编辑按当前位置算;公式引用的自动重写留作后续增强。
 */
import type { CellModel, ColumnInfo, ImageAnchor, MergeRange, RowInfo, SheetModel } from './types'
import { cellKey } from './types'
import { cloneCell } from './snapshot'
import { cloneImageAnchor } from './mutations'

export type StructOp = 'insert-rows' | 'delete-rows' | 'insert-cols' | 'delete-cols'

/** 整张表结构状态快照(供撤销整体还原) */
export interface StructSnapshot {
  cells: Map<string, CellModel>
  merges: MergeRange[]
  rows: Map<number, RowInfo>
  columns: Map<number, ColumnInfo>
  images: ImageAnchor[]
  dimension: { rows: number; cols: number }
}

/** 拍快照(深克隆;images 浅克隆共享 bytes)。 */
export function captureStructure(sheet: SheetModel): StructSnapshot {
  return {
    cells: new Map([...sheet.cells].map(([k, c]) => [k, cloneCell(c)])),
    merges: sheet.merges.map((m) => ({ ...m })),
    rows: new Map([...sheet.rows].map(([k, v]) => [k, { ...v }])),
    columns: new Map([...sheet.columns].map(([k, v]) => [k, { ...v }])),
    images: sheet.images.map(cloneImageAnchor),
    dimension: { ...sheet.dimension },
  }
}

/** 把快照就地还原进 sheet(替换其集合引用,保 sheet 对象身份;再克隆一份保快照可重复用)。 */
export function restoreStructure(sheet: SheetModel, snap: StructSnapshot): void {
  sheet.cells = new Map([...snap.cells].map(([k, c]) => [k, cloneCell(c)]))
  sheet.merges = snap.merges.map((m) => ({ ...m }))
  sheet.rows = new Map([...snap.rows].map(([k, v]) => [k, { ...v }]))
  sheet.columns = new Map([...snap.columns].map(([k, v]) => [k, { ...v }]))
  sheet.images = snap.images.map(cloneImageAnchor)
  sheet.dimension = { ...snap.dimension }
}

/** 把 Map<number,T> 的键做插入/删除位移(插入:>=at 加 count;删除:删 [at,at+count) 段、之后键减 count)。 */
function shiftKeyedMap<T>(map: Map<number, T>, at: number, count: number, mode: 'insert' | 'delete'): Map<number, T> {
  const next = new Map<number, T>()
  for (const [k, v] of map) {
    if (mode === 'insert') {
      next.set(k >= at ? k + count : k, v)
    } else {
      if (k >= at && k < at + count) continue // 删除段
      next.set(k >= at + count ? k - count : k, v)
    }
  }
  return next
}

export function insertRows(sheet: SheetModel, at: number, count: number): void {
  if (count <= 0) return
  const next = new Map<string, CellModel>()
  for (const cell of sheet.cells.values()) {
    if (cell.row >= at) cell.row += count
    next.set(cellKey(cell.row, cell.col), cell)
  }
  sheet.cells = next
  for (const m of sheet.merges) {
    if (m.top >= at) {
      m.top += count
      m.bottom += count
    } else if (m.bottom >= at) m.bottom += count // 跨插入点 → 扩展
  }
  sheet.rows = shiftKeyedMap(sheet.rows, at, count, 'insert')
  for (const im of sheet.images) {
    if (im.from.row >= at) im.from.row += count
    if (im.to && im.to.row >= at) im.to.row += count
  }
  sheet.dimension.rows += count
}

export function deleteRows(sheet: SheetModel, at: number, count: number): void {
  if (count <= 0) return
  const end = at + count
  const next = new Map<string, CellModel>()
  for (const cell of sheet.cells.values()) {
    if (cell.row >= at && cell.row < end) continue // 删除段
    if (cell.row >= end) cell.row -= count
    next.set(cellKey(cell.row, cell.col), cell)
  }
  sheet.cells = next
  sheet.merges = sheet.merges.filter((m) => {
    if (m.bottom < at) return true
    if (m.top >= end) {
      m.top -= count
      m.bottom -= count
      return true
    }
    return false // 与删除段相交 → 丢弃合并(v1 安全策略)
  })
  sheet.rows = shiftKeyedMap(sheet.rows, at, count, 'delete')
  for (const im of sheet.images) {
    if (im.from.row >= end) im.from.row -= count
    else if (im.from.row >= at) im.from.row = at // 落在删除段 → 夹到 at
    if (im.to) {
      if (im.to.row >= end) im.to.row -= count
      else if (im.to.row >= at) im.to.row = at
    }
  }
  sheet.dimension.rows = Math.max(0, sheet.dimension.rows - count)
}

export function insertCols(sheet: SheetModel, at: number, count: number): void {
  if (count <= 0) return
  const next = new Map<string, CellModel>()
  for (const cell of sheet.cells.values()) {
    if (cell.col >= at) cell.col += count
    next.set(cellKey(cell.row, cell.col), cell)
  }
  sheet.cells = next
  for (const m of sheet.merges) {
    if (m.left >= at) {
      m.left += count
      m.right += count
    } else if (m.right >= at) m.right += count
  }
  sheet.columns = shiftKeyedMap(sheet.columns, at, count, 'insert')
  for (const im of sheet.images) {
    if (im.from.col >= at) im.from.col += count
    if (im.to && im.to.col >= at) im.to.col += count
  }
  sheet.dimension.cols += count
}

export function deleteCols(sheet: SheetModel, at: number, count: number): void {
  if (count <= 0) return
  const end = at + count
  const next = new Map<string, CellModel>()
  for (const cell of sheet.cells.values()) {
    if (cell.col >= at && cell.col < end) continue
    if (cell.col >= end) cell.col -= count
    next.set(cellKey(cell.row, cell.col), cell)
  }
  sheet.cells = next
  sheet.merges = sheet.merges.filter((m) => {
    if (m.right < at) return true
    if (m.left >= end) {
      m.left -= count
      m.right -= count
      return true
    }
    return false
  })
  sheet.columns = shiftKeyedMap(sheet.columns, at, count, 'delete')
  for (const im of sheet.images) {
    if (im.from.col >= end) im.from.col -= count
    else if (im.from.col >= at) im.from.col = at
    if (im.to) {
      if (im.to.col >= end) im.to.col -= count
      else if (im.to.col >= at) im.to.col = at
    }
  }
  sheet.dimension.cols = Math.max(0, sheet.dimension.cols - count)
}

/** 应用一个结构操作(分发)。 */
export function applyStructOp(sheet: SheetModel, op: StructOp, at: number, count: number): void {
  if (op === 'insert-rows') insertRows(sheet, at, count)
  else if (op === 'delete-rows') deleteRows(sheet, at, count)
  else if (op === 'insert-cols') insertCols(sheet, at, count)
  else deleteCols(sheet, at, count)
}

/** 删除段是否与任一合并区相交(controller 据此给"会丢合并"警告,仿 sortColumn 守卫)。 */
export function deleteIntersectsMerge(sheet: SheetModel, op: 'delete-rows' | 'delete-cols', at: number, count: number): boolean {
  const end = at + count
  return sheet.merges.some((m) =>
    op === 'delete-rows' ? m.top < end && m.bottom >= at && (m.top < at || m.bottom >= end) : m.left < end && m.right >= at && (m.left < at || m.right >= end),
  )
}
