import { describe, it, expect } from 'vitest'
import { GridMetrics, colIndexToLetters } from '../grid-metrics'
import { makeSheet } from '../../__tests__/helpers'

describe('colIndexToLetters', () => {
  it('单字母 / 进位 / 多字母', () => {
    expect(colIndexToLetters(0)).toBe('A')
    expect(colIndexToLetters(25)).toBe('Z')
    expect(colIndexToLetters(26)).toBe('AA')
    expect(colIndexToLetters(27)).toBe('AB')
    expect(colIndexToLetters(701)).toBe('ZZ')
    expect(colIndexToLetters(702)).toBe('AAA')
  })
})

describe('GridMetrics 几何', () => {
  const m = new GridMetrics(makeSheet({ dimension: { rows: 10, cols: 5 } }))

  it('列左边缘累计(默认列宽 64)', () => {
    expect(m.colLeft(0)).toBe(0)
    expect(m.colLeft(1)).toBe(64)
    expect(m.colLeft(5)).toBe(320) // 数据右边界
  })

  it('超出数据范围按默认列宽外推', () => {
    expect(m.colLeft(7)).toBe(320 + 2 * 64) // 模拟无限网格
  })

  it('坐标→列(含外推)', () => {
    expect(m.colAt(0)).toBe(0)
    expect(m.colAt(70)).toBe(1)
    expect(m.colAt(319)).toBe(4)
    expect(m.colAt(400)).toBe(5 + Math.floor((400 - 320) / 64)) // 数据外
  })

  it('行高换算与坐标→行(默认 20)', () => {
    expect(m.rowTop(3)).toBe(60)
    expect(m.rowAt(50)).toBe(2)
  })

  it('网格线区间可超出数据范围(铺满视口)', () => {
    const [c0, c1] = m.gridColRange(0, 1000)
    expect(c0).toBe(0)
    expect(c1).toBeGreaterThan(5) // 超过数据列,铺空网格
  })

  it('缩放: 几何按 zoom 放大', () => {
    const z = new GridMetrics(makeSheet(), 2)
    expect(z.colLeft(1)).toBe(128)
    expect(z.rowTop(1)).toBe(40)
  })
})
