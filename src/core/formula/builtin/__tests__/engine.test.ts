import { describe, it, expect } from 'vitest'
import { BuiltinFormulaEngine } from '../index'
import { cellKey, type SheetModel, type WorkbookModel } from '../../../model/types'

/** A1 串 → {row,col} */
function a1(ref: string): { row: number; col: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!
  let col = 0
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { row: parseInt(m[2], 10) - 1, col: col - 1 }
}
/** 从 { A1: 值或'=公式' } 造一张 sheet */
function sheet(name: string, cells: Record<string, number | string>): SheetModel {
  const map = new Map()
  let rows = 1, cols = 1
  for (const [ref, val] of Object.entries(cells)) {
    const { row, col } = a1(ref)
    rows = Math.max(rows, row + 1); cols = Math.max(cols, col + 1)
    if (typeof val === 'string' && val[0] === '=') map.set(cellKey(row, col), { row, col, type: 'formula', raw: null, formula: val.slice(1), styleId: 0 })
    else map.set(cellKey(row, col), { row, col, type: typeof val === 'number' ? 'number' : 'string', raw: val, styleId: 0 })
  }
  return { name, index: 0, state: 'visible', dimension: { rows, cols }, cells: map, styles: [], merges: [], columns: new Map(), rows: new Map(), defaultColWidth: 64, defaultRowHeight: 20, freeze: { frozenRows: 0, frozenCols: 0 }, conditional: [], dataValidations: [], images: [], charts: [], shapes: [], sparklines: [], pivotTables: [] } as unknown as SheetModel
}
function wb(...sheets: SheetModel[]): WorkbookModel { return { sheets, date1904: false } as unknown as WorkbookModel }
const val = (e: BuiltinFormulaEngine, ref: string, s = 0) => { const { row, col } = a1(ref); return e.getValue(s, row, col) }
const set = (e: BuiltinFormulaEngine, ref: string, content: string | number | null, s = 0) => { const { row, col } = a1(ref); return e.setCell(s, row, col, content) }

describe('内置公式引擎 — 级联重算', () => {
  it('初始全量求值:=A1+A2', () => {
    const e = new BuiltinFormulaEngine()
    e.setSheets(wb(sheet('S', { A1: 1, A2: 2, A3: '=A1+A2' })))
    expect(val(e, 'A3')).toBe(3)
  })
  it('改前驱 → 依赖格级联更新', () => {
    const e = new BuiltinFormulaEngine()
    e.setSheets(wb(sheet('S', { A1: 1, A2: 2, A3: '=A1+A2' })))
    const dirty = set(e, 'A1', 10)
    expect(val(e, 'A3')).toBe(12)
    expect(dirty.some((d) => d.row === 2 && d.col === 0 && d.value === 12)).toBe(true) // A3 在脏列表
  })
  it('多级链 A1→B1→C1 传递级联', () => {
    const e = new BuiltinFormulaEngine()
    e.setSheets(wb(sheet('S', { A1: 1, B1: '=A1*2', C1: '=B1+1' })))
    expect(val(e, 'C1')).toBe(3)
    set(e, 'A1', 5)
    expect(val(e, 'B1')).toBe(10)
    expect(val(e, 'C1')).toBe(11)
  })
  it('区域 SUM 级联', () => {
    const e = new BuiltinFormulaEngine()
    e.setSheets(wb(sheet('S', { A1: 1, A2: 2, A3: 3, A4: '=SUM(A1:A3)' })))
    expect(val(e, 'A4')).toBe(6)
    set(e, 'A2', 20)
    expect(val(e, 'A4')).toBe(24)
  })
  it('改公式本身', () => {
    const e = new BuiltinFormulaEngine()
    e.setSheets(wb(sheet('S', { A1: 10, A2: '=A1*2' })))
    expect(val(e, 'A2')).toBe(20)
    set(e, 'A2', '=A1+100')
    expect(val(e, 'A2')).toBe(110)
  })
  it('循环引用 → #REF!', () => {
    const e = new BuiltinFormulaEngine()
    e.setSheets(wb(sheet('S', { A1: '=A2', A2: '=A1+1' })))
    expect(val(e, 'A1')).toBe('#REF!')
    expect(val(e, 'A2')).toBe('#REF!')
  })
  it('跨表引用', () => {
    const e = new BuiltinFormulaEngine()
    e.setSheets(wb(sheet('S1', { A1: '=S2!A1+1' }), sheet('S2', { A1: 41 })))
    expect(val(e, 'A1', 0)).toBe(42)
    set(e, 'A1', 100, 1) // 改 S2!A1
    expect(val(e, 'A1', 0)).toBe(101) // S1!A1 跟着变
  })
})
