/**
 * 单元格是否可编辑的判定(纯函数,框架无关)。
 *
 * **优先级**(任一不满足即只读):
 *   ① 总开关 `editable` 关 → 只读
 *   ② `editableTargets` 是**白名单**:**显式设了**(数组或单对象,即使 `[]`)就启用白名单语义,
 *      未命中**任一** target 的格直接只读;**未设(undefined)** = 不启用白名单(默认全可编辑)
 *   ③ 命中 `readOnlyRanges` → 只读(即便在白名单内也能再黑)
 *   ④ `cellReadOnly` 返 truthy → 只读
 *   ⑤ 否则可编辑
 */
import type { MergeRange, SheetModel } from '../model/types'
import { cellKey } from '../model/types'
import type { EditableTarget, EditConfig } from './types'

function inRange(row: number, col: number, r: MergeRange): boolean {
  return row >= r.top && row <= r.bottom && col >= r.left && col <= r.right
}

/** 单格 / 行 / 列 / 区域 四态 target 命中判定. 形状由有哪些字段自动识别. */
export function matchesEditableTarget(row: number, col: number, t: EditableTarget): boolean {
  // 矩形: 有 top/left/bottom/right 4 字段
  if ('top' in t && 'left' in t && 'bottom' in t && 'right' in t) {
    return inRange(row, col, t as MergeRange)
  }
  const hasRow = 'row' in t && typeof (t as { row?: number }).row === 'number'
  const hasCol = 'col' in t && typeof (t as { col?: number }).col === 'number'
  // 单格: row + col
  if (hasRow && hasCol) {
    const tt = t as { row: number; col: number }
    return row === tt.row && col === tt.col
  }
  // 整行: 只有 row
  if (hasRow) return row === (t as { row: number }).row
  // 整列: 只有 col
  if (hasCol) return col === (t as { col: number }).col
  return false
}

/** 把一组格按"是否可编辑"二分. 用于粘贴 / 图片转换 等"批量目标 + 部分跳过"场景. */
export function partitionByEditable(
  sheet: SheetModel,
  cells: Array<{ row: number; col: number }>,
  cfg: EditConfig,
): { allowed: Array<{ row: number; col: number }>; denied: Array<{ row: number; col: number }> } {
  const allowed: Array<{ row: number; col: number }> = []
  const denied: Array<{ row: number; col: number }> = []
  for (const c of cells) {
    if (resolveEditable(sheet, c.row, c.col, cfg)) allowed.push(c)
    else denied.push(c)
  }
  return { allowed, denied }
}

/** 区域是否**全可编辑**(任一格只读即返 ok=false + firstDenied). 用于 mergeCells / unmergeCells. */
export function rangeAllEditable(
  sheet: SheetModel,
  range: MergeRange,
  cfg: EditConfig,
): { ok: boolean; firstDenied?: { row: number; col: number } } {
  for (let r = range.top; r <= range.bottom; r++) {
    for (let c = range.left; c <= range.right; c++) {
      if (!resolveEditable(sheet, r, c, cfg)) return { ok: false, firstDenied: { row: r, col: c } }
    }
  }
  return { ok: true }
}

/** 收集区域里全部不可编辑的格(用于 emit permission-denied 时填 cells 列表). */
export function collectDeniedInRange(
  sheet: SheetModel,
  range: MergeRange,
  cfg: EditConfig,
): Array<{ row: number; col: number }> {
  const out: Array<{ row: number; col: number }> = []
  for (let r = range.top; r <= range.bottom; r++) {
    for (let c = range.left; c <= range.right; c++) {
      if (!resolveEditable(sheet, r, c, cfg)) out.push({ row: r, col: c })
    }
  }
  return out
}

export function resolveEditable(sheet: SheetModel, row: number, col: number, cfg: EditConfig): boolean {
  if (!cfg.editable) return false
  // ② 白名单 —— 显式设了就启用(即便 [] 也启用 = 全只读;undefined 才走"默认全可编辑")
  if (cfg.editableTargets !== undefined) {
    const arr = Array.isArray(cfg.editableTargets) ? cfg.editableTargets : [cfg.editableTargets]
    let hit = false
    for (const t of arr) {
      if (matchesEditableTarget(row, col, t)) { hit = true; break }
    }
    if (!hit) return false
  }
  if (cfg.readOnlyRanges) {
    for (const r of cfg.readOnlyRanges) if (inRange(row, col, r)) return false
  }
  if (cfg.cellReadOnly) {
    const cell = sheet.cells.get(cellKey(row, col)) ?? null
    if (cfg.cellReadOnly(cell, { row, col })) return false
  }
  return true
}
