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
