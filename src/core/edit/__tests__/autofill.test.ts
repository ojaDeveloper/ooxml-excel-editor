import { describe, it, expect } from 'vitest'
import { computeFillSeries } from '../autofill'

describe('自动填充序列引擎 computeFillSeries', () => {
  it('单个数值 → 复制', () => {
    expect(computeFillSeries([5], 3)).toEqual([5, 5, 5])
  })
  it('两个数值 → 等差外推', () => {
    expect(computeFillSeries([1, 2], 3)).toEqual([3, 4, 5])
    expect(computeFillSeries([2, 4], 3)).toEqual([6, 8, 10])
    expect(computeFillSeries([10, 7], 2)).toEqual([4, 1]) // 递减
  })
  it('单个日期 → 每格 +1 天', () => {
    const d = new Date(Date.UTC(2026, 0, 1))
    const out = computeFillSeries([d], 2) as Date[]
    expect(out[0].getTime()).toBe(Date.UTC(2026, 0, 2))
    expect(out[1].getTime()).toBe(Date.UTC(2026, 0, 3))
  })
  it('两个日期 → 按相邻差外推(隔 2 天)', () => {
    const a = new Date(Date.UTC(2026, 0, 1)), b = new Date(Date.UTC(2026, 0, 3))
    const out = computeFillSeries([a, b], 2) as Date[]
    expect(out[0].getTime()).toBe(Date.UTC(2026, 0, 5))
    expect(out[1].getTime()).toBe(Date.UTC(2026, 0, 7))
  })
  it('前缀+末尾整数 → 递增,保留前缀', () => {
    expect(computeFillSeries(['Item 1'], 2)).toEqual(['Item 2', 'Item 3'])
    expect(computeFillSeries(['第1周'], 2)).toEqual(['第2周', '第3周'])
    expect(computeFillSeries(['A01'], 2)).toEqual(['A02', 'A03']) // 前导零位宽保留
  })
  it('星期(中) → 循环接续', () => {
    expect(computeFillSeries(['周五'], 3)).toEqual(['周六', '周日', '周一'])
  })
  it('月份(英缩写) → 循环接续', () => {
    expect(computeFillSeries(['Nov'], 3)).toEqual(['Dec', 'Jan', 'Feb'])
  })
  it('普通文本 → 循环复制', () => {
    expect(computeFillSeries(['甲', '乙'], 3)).toEqual(['甲', '乙', '甲'])
    expect(computeFillSeries(['x'], 2)).toEqual(['x', 'x'])
  })
  it('count<=0 → 空', () => {
    expect(computeFillSeries([1], 0)).toEqual([])
  })

  // ---- Ctrl 翻转 复制↔序列(对齐 Excel)----
  it('Ctrl + 单个数值 → 递增(+1)', () => {
    expect(computeFillSeries([5], 3, true)).toEqual([6, 7, 8])
  })
  it('Ctrl + 两个数值序列 → 复制(循环源块)', () => {
    expect(computeFillSeries([1, 2], 3, true)).toEqual([1, 2, 1])
  })
  it('Ctrl + 单个日期 → 复制(不再 +1 天)', () => {
    const d = new Date(Date.UTC(2026, 0, 1))
    expect((computeFillSeries([d], 2, true) as Date[]).map((x) => x.getTime())).toEqual([d.getTime(), d.getTime()])
  })
  it('Ctrl + 文本递增/星期 → 复制', () => {
    expect(computeFillSeries(['Item 1'], 2, true)).toEqual(['Item 1', 'Item 1'])
    expect(computeFillSeries(['周五'], 2, true)).toEqual(['周五', '周五'])
  })
})
