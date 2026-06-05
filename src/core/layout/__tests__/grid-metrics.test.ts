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

describe('GridMetrics 虚拟范围(滚动出空行/列)', () => {
  const sheet = makeSheet({ dimension: { rows: 10, cols: 5 } }) // 默认 64×20
  it('virtualWidth/Height 含外推,而 totalWidth/Height 仍按 dimension 不变', () => {
    const m = new GridMetrics(sheet, 1, 100, 20) // 虚拟 100 行 / 20 列
    expect(m.totalWidth).toBe(320) // 5 列 × 64,不受虚拟影响(导出/数据安全)
    expect(m.totalHeight).toBe(200) // 10 行 × 20
    expect(m.virtualWidth).toBe(20 * 64) // 20 虚拟列
    expect(m.virtualHeight).toBe(100 * 20) // 100 虚拟行
    expect(m.vRows).toBe(100)
    expect(m.vCols).toBe(20)
  })
  it('虚拟数 < dimension 时取 dimension(只增不减)', () => {
    const m = new GridMetrics(sheet, 1, 3, 2)
    expect(m.vRows).toBe(10) // max(dim 10, 虚拟 3)
    expect(m.vCols).toBe(5)
  })
  it('封顶 Excel 上限', () => {
    const m = new GridMetrics(sheet, 1, 9_999_999, 99_999)
    expect(m.vRows).toBe(1048576)
    expect(m.vCols).toBe(16384)
  })
  it('可视区范围夹到虚拟范围(允许画/选空行)', () => {
    const m = new GridMetrics(sheet, 1, 100, 20)
    const [, r1] = m.visibleRowRange(0, 100000) // 视口很高
    expect(r1).toBe(99) // vRows-1,而非 dimension 的 9
  })
})
