import { describe, it, expect } from 'vitest'
import { definePlugin } from '../plugin'

describe('definePlugin', () => {
  it('原样返回插件对象(仅类型推断)', () => {
    const p = definePlugin({
      name: 'x',
      theme: { gridLine: '#f00' },
      cellStyle: (c) => (typeof c.raw === 'number' && c.raw < 0 ? { font: { color: '#d00' } } : undefined),
      events: { 'cell-click': () => {} },
    })
    expect(p.name).toBe('x')
    expect(p.theme?.gridLine).toBe('#f00')
    expect(typeof p.cellStyle).toBe('function')
    expect(typeof p.events?.['cell-click']).toBe('function')
  })

  it('cellStyle 钩子按值返回部分样式', () => {
    const p = definePlugin({
      name: 'neg',
      cellStyle: (c) => (typeof c.raw === 'number' && c.raw < 0 ? { font: { color: '#d00', bold: true } } : undefined),
    })
    const neg = p.cellStyle!({ row: 0, col: 0, type: 'number', raw: -5, styleId: 0 }, { row: 0, col: 0 })
    const pos = p.cellStyle!({ row: 0, col: 0, type: 'number', raw: 5, styleId: 0 }, { row: 0, col: 0 })
    expect((neg as any)?.font?.color).toBe('#d00')
    expect(pos).toBeUndefined()
  })
})
