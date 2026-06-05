import { describe, it, expect } from 'vitest'
import { defaultFormulaEngineFactory } from '../hyperformula-adapter'
import { cellContentForEngine } from '../engine'
import type { CellModel, SheetModel, WorkbookModel } from '../../model/types'
import { cellKey } from '../../model/types'

function cell(row: number, col: number, type: CellModel['type'], raw: CellModel['raw'], formula?: string): CellModel {
  return { row, col, type, raw, formula, styleId: 0 }
}
function sheet(name: string, cells: CellModel[], rows: number, cols: number): SheetModel {
  const map = new Map<string, CellModel>()
  for (const c of cells) map.set(cellKey(c.row, c.col), c)
  return { name, cells: map, dimension: { rows, cols } } as unknown as SheetModel
}

describe('cellContentForEngine(喂引擎的内容)', () => {
  it('公式→公式串;值→字面;空/日期→null', () => {
    expect(cellContentForEngine(cell(0, 0, 'formula', 6, '=A1+1'))).toBe('=A1+1')
    expect(cellContentForEngine(cell(0, 0, 'formula', 6, 'B3*C3'))).toBe('=B3*C3') // 解析层不带 '=' → 补上
    expect(cellContentForEngine(cell(0, 0, 'number', 42))).toBe(42)
    expect(cellContentForEngine(cell(0, 0, 'string', 'hi'))).toBe('hi')
    expect(cellContentForEngine(cell(0, 0, 'boolean', true))).toBe(true)
    expect(cellContentForEngine(cell(0, 0, 'date', new Date()))).toBeNull()
    expect(cellContentForEngine(null)).toBeNull()
  })
})

describe('HyperFormulaAdapter(真引擎冒烟:地址映射 / 级联 / 跨表 / 错误)', () => {
  it('setCell 级联 + 0-based 地址往返 + 错误映射', async () => {
    const engine = await defaultFormulaEngineFactory()
    // A1=5 B1=10 C1==A1+B1(15);A2=2 B2=3 C2==A2*B2(6)
    const s1 = sheet(
      'S1',
      [
        cell(0, 0, 'number', 5),
        cell(0, 1, 'number', 10),
        cell(0, 2, 'formula', 15, '=A1+B1'),
        cell(1, 0, 'number', 2),
        cell(1, 1, 'number', 3),
        cell(1, 2, 'formula', 6, '=A2*B2'),
      ],
      2,
      3,
    )
    const wb = { sheets: [s1], activeSheet: 0, themeColors: [], date1904: false } as unknown as WorkbookModel
    engine.setSheets(wb)

    expect(engine.getValue(0, 0, 2)).toBe(15) // C1 初值
    // 改 A1=100 → C1 级联到 110(返回脏格含 A1 + C1)
    const dirty = engine.setCell(0, 0, 0, 100)
    const c1 = dirty.find((d) => d.row === 0 && d.col === 2)
    expect(c1?.value).toBe(110)
    expect(engine.getValue(0, 0, 2)).toBe(110)

    // 错误映射:C1 = 1/0 → '#DIV/0!' 串
    engine.setCell(0, 0, 2, '=1/0')
    expect(engine.getValue(0, 0, 2)).toBe('#DIV/0!')
    engine.destroy()
  })

  it('跨表引用:S2 公式引用 S1 的格', async () => {
    const engine = await defaultFormulaEngineFactory()
    const s1 = sheet('S1', [cell(0, 0, 'number', 7)], 1, 1)
    const s2 = sheet('S2', [cell(0, 0, 'formula', 7, '=S1!A1*2')], 1, 1)
    const wb = { sheets: [s1, s2], activeSheet: 0, themeColors: [], date1904: false } as unknown as WorkbookModel
    engine.setSheets(wb)
    expect(engine.getValue(1, 0, 0)).toBe(14) // S2!A1 = S1!A1*2
    // 改 S1!A1=10 → S2!A1 级联到 20(脏格含 sheet:1 的格)
    const dirty = engine.setCell(0, 0, 0, 10)
    const s2dep = dirty.find((d) => d.sheet === 1 && d.row === 0 && d.col === 0)
    expect(s2dep?.value).toBe(20)
    engine.destroy()
  })
})
