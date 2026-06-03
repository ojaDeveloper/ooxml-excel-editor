import { describe, it, expect } from 'vitest'
import { ConditionalEngine } from '../conditional'
import type { ConditionalRule } from '../../model/types'
import { makeSheet, putNumbers } from '../../__tests__/helpers'

const colA = { top: 0, left: 0, bottom: 4, right: 0 } // A1:A5

function sheetWith(rule: ConditionalRule) {
  const s = makeSheet({ conditional: [rule] })
  putNumbers(s, [
    { row: 0, col: 0, v: 1 },
    { row: 1, col: 0, v: 2 },
    { row: 2, col: 0, v: 3 },
    { row: 3, col: 0, v: 4 },
    { row: 4, col: 0, v: 5 },
  ])
  return s
}

describe('ConditionalEngine', () => {
  it('colorScale: 两端插值到 min/max 颜色', () => {
    const eng = new ConditionalEngine(
      sheetWith({ ranges: [colA], priority: 1, type: 'colorScale', colorScale: { min: '#FF0000', max: '#00FF00' } }),
    )
    expect(eng.effectsFor(0, 0, 1)?.fillColor).toBe('rgb(255,0,0)') // 最小值 → 红
    expect(eng.effectsFor(4, 0, 5)?.fillColor).toBe('rgb(0,255,0)') // 最大值 → 绿
  })

  it('dataBar: ratio 随值线性变化', () => {
    const eng = new ConditionalEngine(
      sheetWith({ ranges: [colA], priority: 1, type: 'dataBar', dataBar: { color: '#638EC6', gradient: true } }),
    )
    expect(eng.effectsFor(4, 0, 5)?.dataBar?.ratio).toBeCloseTo(1, 5)
    expect(eng.effectsFor(0, 0, 1)?.dataBar?.ratio).toBeCloseTo(0.2, 5)
  })

  it('cellIs greaterThan: 命中套样式,未命中无效果', () => {
    const eng = new ConditionalEngine(
      sheetWith({
        ranges: [colA],
        priority: 1,
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['3'],
        style: { fill: { type: 'solid', fgColor: '#FFFF00' } },
      }),
    )
    expect(eng.effectsFor(4, 0, 5)?.fillColor).toBe('#FFFF00') // 5 > 3
    expect(eng.effectsFor(0, 0, 1)).toBeNull() // 1 > 3 不成立
  })

  it('hasRules 反映是否有规则', () => {
    expect(new ConditionalEngine(makeSheet()).hasRules()).toBe(false)
  })
})
