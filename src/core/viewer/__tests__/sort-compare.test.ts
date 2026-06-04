import { describe, it, expect } from 'vitest'
import { compareCellValues, isBlankValue } from '../controller'

/** 排序比较器:类型序(数字/日期 < 文本 < 布尔)+ 同类型内自然比较;空值由调用方排末尾。 */
describe('compareCellValues / isBlankValue(排序比较)', () => {
  it('空值判定', () => {
    expect(isBlankValue(null)).toBe(true)
    expect(isBlankValue('')).toBe(true)
    expect(isBlankValue(0)).toBe(false)
    expect(isBlankValue('a')).toBe(false)
    expect(isBlankValue(false)).toBe(false)
  })

  it('数字按大小', () => {
    expect(compareCellValues(1, 2)).toBeLessThan(0)
    expect(compareCellValues(10, 2)).toBeGreaterThan(0)
    expect(compareCellValues(3, 3)).toBe(0)
    expect(compareCellValues(-5, 1)).toBeLessThan(0)
  })

  it('文本自然顺序(数字混排)', () => {
    expect(compareCellValues('a2', 'a10')).toBeLessThan(0) // 自然序:2 < 10
    expect(compareCellValues('苹果', '香蕉')).toBeLessThan(0) // 拼音 p < x
  })

  it('类型序:数字 < 文本 < 布尔', () => {
    expect(compareCellValues(100, 'abc')).toBeLessThan(0)
    expect(compareCellValues('abc', true)).toBeLessThan(0)
    expect(compareCellValues(true, 100)).toBeGreaterThan(0)
  })

  it('日期按时间先后(并与数字同档比较)', () => {
    expect(compareCellValues(new Date('2020-01-01'), new Date('2021-01-01'))).toBeLessThan(0)
    expect(compareCellValues(new Date('2020-06-01'), new Date('2020-01-01'))).toBeGreaterThan(0)
  })

  it('布尔 false < true', () => {
    expect(compareCellValues(false, true)).toBeLessThan(0)
    expect(compareCellValues(true, true)).toBe(0)
  })
})
