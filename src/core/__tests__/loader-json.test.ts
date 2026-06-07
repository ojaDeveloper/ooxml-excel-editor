import { describe, it, expect } from 'vitest'
import { jsonToWorkbook, isWorkbookModel } from '../loader-json'
import { cellKey } from '../model/types'

describe('jsonToWorkbook — 三种 input shape', () => {
  it('二维数组:单表,首格 A1', () => {
    const wb = jsonToWorkbook([
      [1, 'hi', true],
      [2, 'bye', false],
    ])
    expect(wb.sheets).toHaveLength(1)
    const s = wb.sheets[0]
    expect(s.name).toBe('Sheet1')
    expect(s.dimension).toEqual({ rows: 2, cols: 3 })
    expect(s.cells.get(cellKey(0, 0))?.raw).toBe(1)
    expect(s.cells.get(cellKey(0, 1))?.raw).toBe('hi')
    expect(s.cells.get(cellKey(0, 2))?.raw).toBe(true)
    expect(s.cells.get(cellKey(1, 0))?.raw).toBe(2)
  })

  it('对象数组:首行写表头 = keys,数据从第二行起', () => {
    const wb = jsonToWorkbook([
      { name: 'Alice', age: 30 },
      { name: 'Bob', age: 25 },
    ])
    const s = wb.sheets[0]
    expect(s.cells.get(cellKey(0, 0))?.raw).toBe('name')
    expect(s.cells.get(cellKey(0, 1))?.raw).toBe('age')
    expect(s.cells.get(cellKey(1, 0))?.raw).toBe('Alice')
    expect(s.cells.get(cellKey(1, 1))?.raw).toBe(30)
    expect(s.dimension).toEqual({ rows: 3, cols: 2 })
  })

  it('对象数组 + headerRow=false:不写表头', () => {
    const wb = jsonToWorkbook([{ a: 1, b: 2 }], { headerRow: false })
    const s = wb.sheets[0]
    expect(s.cells.get(cellKey(0, 0))?.raw).toBe(1)
    expect(s.cells.get(cellKey(0, 1))?.raw).toBe(2)
  })

  it('{ sheets: [...] } 多表 + 自定义名', () => {
    const wb = jsonToWorkbook({
      sheets: [
        { name: 'A', rows: [[1, 2]] },
        { name: 'B', rows: [{ x: 10, y: 20 }] },
      ],
    })
    expect(wb.sheets).toHaveLength(2)
    expect(wb.sheets[0].name).toBe('A')
    expect(wb.sheets[0].cells.get(cellKey(0, 0))?.raw).toBe(1)
    expect(wb.sheets[1].name).toBe('B')
    // B 表对象数组:第 0 行是表头 x/y,第 1 行是 10/20
    expect(wb.sheets[1].cells.get(cellKey(0, 0))?.raw).toBe('x')
    expect(wb.sheets[1].cells.get(cellKey(1, 0))?.raw).toBe(10)
  })

  it('类型推断:数字字符串→number / TRUE/FALSE→boolean / ISO 日期串→Date', () => {
    const wb = jsonToWorkbook([['42', 'TRUE', '2026-01-01', 'plain text']])
    const s = wb.sheets[0]
    expect(s.cells.get(cellKey(0, 0))?.raw).toBe(42)
    expect(s.cells.get(cellKey(0, 0))?.type).toBe('number')
    expect(s.cells.get(cellKey(0, 1))?.raw).toBe(true)
    expect(s.cells.get(cellKey(0, 2))?.raw).toBeInstanceOf(Date)
    expect(s.cells.get(cellKey(0, 3))?.raw).toBe('plain text')
  })

  it('autoInfer=false 时纯字符串照常保留', () => {
    const wb = jsonToWorkbook([['42', 'TRUE']], { autoInfer: false })
    const s = wb.sheets[0]
    expect(s.cells.get(cellKey(0, 0))?.raw).toBe('42')
    expect(s.cells.get(cellKey(0, 0))?.type).toBe('string')
    expect(s.cells.get(cellKey(0, 1))?.raw).toBe('TRUE')
  })

  it('空数组 → 空表(dimension=0×0)', () => {
    const wb = jsonToWorkbook([])
    expect(wb.sheets[0].dimension).toEqual({ rows: 0, cols: 0 })
    expect(wb.sheets[0].cells.size).toBe(0)
  })
})

describe('isWorkbookModel', () => {
  it('完整 shape 返 true', () => {
    expect(isWorkbookModel({ sheets: [], activeSheet: 0, themeColors: [], date1904: false })).toBe(true)
  })
  it('缺字段返 false', () => {
    expect(isWorkbookModel({ sheets: [] })).toBe(false)
    expect(isWorkbookModel([[1, 2]])).toBe(false)
    expect(isWorkbookModel(null)).toBe(false)
  })
})
