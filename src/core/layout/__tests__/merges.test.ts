import { describe, it, expect } from 'vitest'
import { MergeIndex } from '../merges'
import { makeSheet } from '../../__tests__/helpers'

describe('MergeIndex', () => {
  const sheet = makeSheet({
    merges: [
      { top: 0, left: 0, bottom: 0, right: 4 }, // A1:E1
      { top: 2, left: 1, bottom: 4, right: 2 }, // B3:C5
    ],
  })
  const idx = new MergeIndex(sheet)

  it('锚点识别', () => {
    expect(idx.isAnchor(0, 0)).toBe(true)
    expect(idx.isAnchor(2, 1)).toBe(true)
    expect(idx.isAnchor(0, 1)).toBe(false) // 被覆盖但非锚
  })

  it('成员都能查到所属合并区', () => {
    expect(idx.rangeOf(0, 3)).toEqual({ top: 0, left: 0, bottom: 0, right: 4 })
    expect(idx.rangeOf(3, 2)).toEqual({ top: 2, left: 1, bottom: 4, right: 2 })
    expect(idx.rangeOf(5, 5)).toBeUndefined() // 区外
  })

  it('被覆盖(非锚)判定', () => {
    expect(idx.isCovered(0, 2)).toBe(true) // A1:E1 内非锚
    expect(idx.isCovered(0, 0)).toBe(false) // 锚点不算被覆盖
    expect(idx.isCovered(9, 9)).toBe(false) // 区外
  })
})
