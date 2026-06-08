import { describe, it, expect } from 'vitest'
import { resolveEditable } from '../permissions'
import type { EditConfig } from '../types'
import type { CellModel, SheetModel } from '../../model/types'
import { cellKey } from '../../model/types'

/** 极简 sheet:resolveEditable 只用到 sheet.cells.get */
function sheetWith(cells: Record<string, Partial<CellModel>> = {}): SheetModel {
  const map = new Map<string, CellModel>()
  for (const [k, v] of Object.entries(cells)) map.set(k, { row: 0, col: 0, type: 'string', raw: '', styleId: 0, ...v })
  return { cells: map } as unknown as SheetModel
}

describe('resolveEditable(编辑权限优先级)', () => {
  it('默认(editable 未开)→ 一律只读', () => {
    const s = sheetWith()
    expect(resolveEditable(s, 0, 0, {})).toBe(false)
    expect(resolveEditable(s, 5, 5, { editable: false })).toBe(false)
  })

  it('editable 开 + 无其它限制 → 可编辑', () => {
    expect(resolveEditable(sheetWith(), 3, 4, { editable: true })).toBe(true)
  })

  it('readOnlyRanges 命中 → 只读(优先于 cellReadOnly)', () => {
    const cfg: EditConfig = { editable: true, readOnlyRanges: [{ top: 1, left: 1, bottom: 3, right: 3 }] }
    expect(resolveEditable(sheetWith(), 2, 2, cfg)).toBe(false) // 命中
    expect(resolveEditable(sheetWith(), 0, 0, cfg)).toBe(true) // 区域外
    expect(resolveEditable(sheetWith(), 3, 3, cfg)).toBe(false) // 闭区间右下角
    expect(resolveEditable(sheetWith(), 4, 4, cfg)).toBe(true) // 区域外
  })

  it('cellReadOnly 返 true → 只读;空格传 null', () => {
    const seen: Array<CellModel | null> = []
    const cfg: EditConfig = {
      editable: true,
      cellReadOnly: (cell, pos) => {
        seen.push(cell)
        return pos.col === 0 // 第 0 列只读
      },
    }
    const s = sheetWith({ [cellKey(0, 0)]: { raw: 'A' } })
    expect(resolveEditable(s, 0, 0, cfg)).toBe(false) // 第 0 列
    expect(resolveEditable(s, 0, 1, cfg)).toBe(true) // 其它列(空格)
    expect(seen[0]).toMatchObject({ raw: 'A' }) // 有格传 CellModel
    expect(seen[1]).toBeNull() // 空格传 null
  })

  it('cellReadOnly 返 falsy(void/false)→ 不只读', () => {
    const cfg: EditConfig = { editable: true, cellReadOnly: () => undefined }
    expect(resolveEditable(sheetWith(), 0, 0, cfg)).toBe(true)
  })
})

describe('resolveEditable — editableTargets 白名单 (2026-06-08)', () => {
  it('未设 (undefined) → 不启用白名单, 老行为 (默认全可编辑)', () => {
    const s = sheetWith()
    expect(resolveEditable(s, 0, 0, { editable: true })).toBe(true)
    expect(resolveEditable(s, 100, 100, { editable: true })).toBe(true)
  })

  it('显式空数组 [] → 全只读 (没格在白名单)', () => {
    const s = sheetWith()
    const cfg: EditConfig = { editable: true, editableTargets: [] }
    expect(resolveEditable(s, 0, 0, cfg)).toBe(false)
    expect(resolveEditable(s, 5, 5, cfg)).toBe(false)
  })

  it('单格 target (单对象, 非数组) → 只该格可编辑', () => {
    const s = sheetWith()
    const cfg: EditConfig = { editable: true, editableTargets: { row: 2, col: 3 } }
    expect(resolveEditable(s, 2, 3, cfg)).toBe(true)
    expect(resolveEditable(s, 2, 4, cfg)).toBe(false)
    expect(resolveEditable(s, 3, 3, cfg)).toBe(false)
  })

  it('多个不相邻单格 → 只那些格可编辑 (典型: 协同编辑)', () => {
    const s = sheetWith()
    const cfg: EditConfig = {
      editable: true,
      editableTargets: [
        { row: 0, col: 0 },
        { row: 5, col: 7 },
        { row: 10, col: 2 },
      ],
    }
    expect(resolveEditable(s, 0, 0, cfg)).toBe(true)
    expect(resolveEditable(s, 5, 7, cfg)).toBe(true)
    expect(resolveEditable(s, 10, 2, cfg)).toBe(true)
    expect(resolveEditable(s, 0, 1, cfg)).toBe(false)
    expect(resolveEditable(s, 5, 8, cfg)).toBe(false)
    expect(resolveEditable(s, 10, 3, cfg)).toBe(false)
  })

  it('整行 target ({row}) → 该行所有列可编辑', () => {
    const s = sheetWith()
    const cfg: EditConfig = { editable: true, editableTargets: { row: 3 } }
    expect(resolveEditable(s, 3, 0, cfg)).toBe(true)
    expect(resolveEditable(s, 3, 100, cfg)).toBe(true)
    expect(resolveEditable(s, 2, 0, cfg)).toBe(false)
    expect(resolveEditable(s, 4, 0, cfg)).toBe(false)
  })

  it('整列 target ({col}) → 该列所有行可编辑', () => {
    const s = sheetWith()
    const cfg: EditConfig = { editable: true, editableTargets: { col: 5 } }
    expect(resolveEditable(s, 0, 5, cfg)).toBe(true)
    expect(resolveEditable(s, 100, 5, cfg)).toBe(true)
    expect(resolveEditable(s, 0, 4, cfg)).toBe(false)
    expect(resolveEditable(s, 0, 6, cfg)).toBe(false)
  })

  it('矩形 target (MergeRange) → 区域内可编辑', () => {
    const s = sheetWith()
    const cfg: EditConfig = {
      editable: true,
      editableTargets: { top: 1, left: 1, bottom: 3, right: 3 },
    }
    expect(resolveEditable(s, 1, 1, cfg)).toBe(true) // 左上角
    expect(resolveEditable(s, 3, 3, cfg)).toBe(true) // 右下角 (闭区间)
    expect(resolveEditable(s, 2, 2, cfg)).toBe(true) // 中间
    expect(resolveEditable(s, 0, 0, cfg)).toBe(false) // 区域外
    expect(resolveEditable(s, 4, 4, cfg)).toBe(false) // 区域外
  })

  it('混合 4 种 target 在一个数组里 → 各自独立命中, 全部支持', () => {
    const s = sheetWith()
    const cfg: EditConfig = {
      editable: true,
      editableTargets: [
        { row: 0, col: 0 }, // 单格
        { row: 2 }, // 整行
        { col: 4 }, // 整列
        { top: 5, left: 5, bottom: 6, right: 6 }, // 矩形
      ],
    }
    expect(resolveEditable(s, 0, 0, cfg)).toBe(true) // 单格命中
    expect(resolveEditable(s, 2, 99, cfg)).toBe(true) // 整行命中
    expect(resolveEditable(s, 99, 4, cfg)).toBe(true) // 整列命中
    expect(resolveEditable(s, 5, 5, cfg)).toBe(true) // 矩形命中
    expect(resolveEditable(s, 6, 6, cfg)).toBe(true) // 矩形右下角
    // 一处也不命中
    expect(resolveEditable(s, 1, 1, cfg)).toBe(false)
    expect(resolveEditable(s, 7, 7, cfg)).toBe(false)
  })

  it('白名单内的格仍能被 readOnlyRanges 黑掉 (优先级 ③: 白名单内 ∩ 黑名单 → 只读)', () => {
    const s = sheetWith()
    const cfg: EditConfig = {
      editable: true,
      editableTargets: [{ top: 0, left: 0, bottom: 5, right: 5 }], // 白名单 6×6
      readOnlyRanges: [{ top: 2, left: 2, bottom: 3, right: 3 }], // 黑名单 (在白名单内挖一块)
    }
    expect(resolveEditable(s, 0, 0, cfg)).toBe(true) // 白名单内, 不在黑名单
    expect(resolveEditable(s, 2, 2, cfg)).toBe(false) // 白名单内, 但被黑名单
    expect(resolveEditable(s, 3, 3, cfg)).toBe(false) // 白名单内, 但被黑名单
    expect(resolveEditable(s, 6, 6, cfg)).toBe(false) // 不在白名单
  })

  it('editable=false 时白名单也无效 (优先级 ①)', () => {
    const s = sheetWith()
    const cfg: EditConfig = { editable: false, editableTargets: [{ row: 0, col: 0 }] }
    expect(resolveEditable(s, 0, 0, cfg)).toBe(false)
  })
})

describe('partitionByEditable / rangeAllEditable / collectDeniedInRange — Phase A helpers (2026-06-08)', () => {
  it('partitionByEditable: 把 5 个格 (3 允许 / 2 拒) 拆开', async () => {
    const { partitionByEditable } = await import('../permissions')
    const s = sheetWith()
    const cfg: EditConfig = {
      editable: true,
      editableTargets: [{ row: 0, col: 0 }, { row: 1, col: 1 }, { row: 2, col: 2 }],
    }
    const cells = [
      { row: 0, col: 0 }, // ✓
      { row: 1, col: 1 }, // ✓
      { row: 2, col: 2 }, // ✓
      { row: 0, col: 1 }, // ✗
      { row: 3, col: 3 }, // ✗
    ]
    const { allowed, denied } = partitionByEditable(s, cells, cfg)
    expect(allowed).toHaveLength(3)
    expect(denied).toHaveLength(2)
    expect(denied[0]).toEqual({ row: 0, col: 1 })
  })

  it('rangeAllEditable: 区域内任一格只读 → ok=false + firstDenied', async () => {
    const { rangeAllEditable } = await import('../permissions')
    const s = sheetWith()
    const cfg: EditConfig = { editable: true, editableTargets: [{ row: 0, col: 0 }] }
    // 区域 (0,0)-(1,1) 含 (0,1),(1,0),(1,1) 三个非白名单格
    const got = rangeAllEditable(s, { top: 0, left: 0, bottom: 1, right: 1 }, cfg)
    expect(got.ok).toBe(false)
    expect(got.firstDenied).toEqual({ row: 0, col: 1 })
  })

  it('rangeAllEditable: 区域全可编辑 → ok=true', async () => {
    const { rangeAllEditable } = await import('../permissions')
    const s = sheetWith()
    const cfg: EditConfig = { editable: true }
    const got = rangeAllEditable(s, { top: 0, left: 0, bottom: 2, right: 2 }, cfg)
    expect(got.ok).toBe(true)
    expect(got.firstDenied).toBeUndefined()
  })

  it('collectDeniedInRange: 返回区域内全部 denied 格', async () => {
    const { collectDeniedInRange } = await import('../permissions')
    const s = sheetWith()
    const cfg: EditConfig = { editable: true, editableTargets: [{ row: 1, col: 1 }] }
    // 3×3 区域只有中心 (1,1) 允许, 其余 8 格 denied
    const denied = collectDeniedInRange(s, { top: 0, left: 0, bottom: 2, right: 2 }, cfg)
    expect(denied).toHaveLength(8)
    // 验证不包含 (1,1)
    expect(denied.some((c) => c.row === 1 && c.col === 1)).toBe(false)
  })
})

describe('matchesEditableTarget — 形状识别', () => {
  it('每种形状自动识别, 互不混淆', async () => {
    const { matchesEditableTarget } = await import('../permissions')
    // 单格
    expect(matchesEditableTarget(1, 2, { row: 1, col: 2 })).toBe(true)
    expect(matchesEditableTarget(1, 3, { row: 1, col: 2 })).toBe(false)
    // 整行: 只有 row, 没 col
    expect(matchesEditableTarget(1, 999, { row: 1 })).toBe(true)
    expect(matchesEditableTarget(2, 999, { row: 1 })).toBe(false)
    // 整列: 只有 col, 没 row
    expect(matchesEditableTarget(999, 4, { col: 4 })).toBe(true)
    expect(matchesEditableTarget(999, 5, { col: 4 })).toBe(false)
    // 矩形
    expect(matchesEditableTarget(2, 2, { top: 1, left: 1, bottom: 3, right: 3 })).toBe(true)
    expect(matchesEditableTarget(4, 4, { top: 1, left: 1, bottom: 3, right: 3 })).toBe(false)
  })
})
