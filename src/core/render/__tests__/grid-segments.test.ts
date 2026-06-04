import { describe, it, expect } from 'vitest'
import { gridSegments } from '../canvas-renderer'

/**
 * 网格线"挖空合并区"的补集逻辑 —— 合并单元格内部不画网格线(跟 Excel/WPS 一致)。
 * gridSegments(start,end,gaps) = [start,end] 减去 gaps 并集后剩下的要画的实线段。
 */
describe('gridSegments(网格线跳过合并区)', () => {
  it('无合并 → 整条线', () => {
    expect(gridSegments(0, 100, [])).toEqual([[0, 100]])
  })

  it('中间一个合并区 → 断成两段', () => {
    expect(gridSegments(0, 100, [[30, 60]])).toEqual([
      [0, 30],
      [60, 100],
    ])
  })

  it('合并区贴着起点 → 只剩后半段(合并外边界处不画该内部线)', () => {
    expect(gridSegments(0, 100, [[0, 40]])).toEqual([[40, 100]])
  })

  it('合并区覆盖整条 → 完全不画', () => {
    expect(gridSegments(0, 100, [[0, 100]])).toEqual([])
    expect(gridSegments(10, 90, [[0, 100]])).toEqual([])
  })

  it('多个合并区(乱序 + 重叠)→ 取并集补集', () => {
    expect(gridSegments(0, 200, [[120, 160], [30, 60], [50, 80]])).toEqual([
      [0, 30],
      [80, 120],
      [160, 200],
    ])
  })

  it('相邻合并区(首尾相接)→ 合并成一个空隙', () => {
    expect(gridSegments(0, 120, [[30, 60], [60, 90]])).toEqual([
      [0, 30],
      [90, 120],
    ])
  })

  it('合并区超出范围 → 裁剪到 [start,end]', () => {
    expect(gridSegments(20, 80, [[0, 40]])).toEqual([[40, 80]])
    expect(gridSegments(20, 80, [[60, 200]])).toEqual([[20, 60]])
  })
})
