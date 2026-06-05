import { describe, it, expect } from 'vitest'
import {
  setCellValue,
  clearCell,
  setRangeValues,
  restoreCell,
  internStyle,
  setColumnWidth,
  setRowHeight,
  restoreDimension,
  applyStyleOverride,
  mergeStyleOverride,
} from '../mutations'
import type { CellStyle, SheetModel } from '../types'
import { cellKey } from '../types'

const styleA = { font: {}, fill: { type: 'none' }, numFmt: 'General' } as unknown as CellStyle
function sheet(): SheetModel {
  return {
    cells: new Map(),
    styles: [styleA],
    columns: new Map(),
    rows: new Map(),
    dimension: { rows: 0, cols: 0 },
  } as unknown as SheetModel
}

describe('mutations.setCellValue(类型推断 + dimension)', () => {
  it('数字 / 布尔 / 日期 / 字符串 / 公式 类型推断', () => {
    const s = sheet()
    setCellValue(s, 0, 0, 42)
    setCellValue(s, 0, 1, true)
    const d = new Date('2026-01-01')
    setCellValue(s, 0, 2, d)
    setCellValue(s, 0, 3, 'hello')
    setCellValue(s, 0, 4, '=A1+1')
    expect(s.cells.get(cellKey(0, 0))).toMatchObject({ type: 'number', raw: 42 })
    expect(s.cells.get(cellKey(0, 1))).toMatchObject({ type: 'boolean', raw: true })
    expect(s.cells.get(cellKey(0, 2))).toMatchObject({ type: 'date' })
    expect(s.cells.get(cellKey(0, 3))).toMatchObject({ type: 'string', raw: 'hello' })
    expect(s.cells.get(cellKey(0, 4))).toMatchObject({ type: 'formula', formula: '=A1+1', raw: null })
  })

  it('纯数字字符串 → 数字;空串 → 删除', () => {
    const s = sheet()
    setCellValue(s, 1, 0, '3.14')
    expect(s.cells.get(cellKey(1, 0))).toMatchObject({ type: 'number', raw: 3.14 })
    setCellValue(s, 1, 0, '')
    expect(s.cells.has(cellKey(1, 0))).toBe(false)
  })

  it('保留原 styleId;dimension 随写入增长', () => {
    const s = sheet()
    s.cells.set(cellKey(2, 2), { row: 2, col: 2, type: 'string', raw: 'x', styleId: 7 } as never)
    s.styles[7] = styleA
    setCellValue(s, 2, 2, 99)
    expect(s.cells.get(cellKey(2, 2))).toMatchObject({ raw: 99, styleId: 7 }) // styleId 保留
    setCellValue(s, 5, 9, 1)
    expect(s.dimension).toMatchObject({ rows: 6, cols: 10 }) // 增长到 row+1/col+1
  })
})

describe('mutations 其它', () => {
  it('clearCell 删除该格', () => {
    const s = sheet()
    setCellValue(s, 0, 0, 1)
    clearCell(s, 0, 0)
    expect(s.cells.has(cellKey(0, 0))).toBe(false)
  })

  it('setRangeValues 左上对齐铺值', () => {
    const s = sheet()
    setRangeValues(s, { top: 1, left: 1, bottom: 2, right: 2 }, [
      [1, 2],
      [3, 4],
    ])
    expect(s.cells.get(cellKey(1, 1))).toMatchObject({ raw: 1 })
    expect(s.cells.get(cellKey(2, 2))).toMatchObject({ raw: 4 })
  })

  it('restoreCell 精确还原(写回 / 删除)', () => {
    const s = sheet()
    const prev = { row: 0, col: 0, type: 'string', raw: 'orig', styleId: 0 } as never
    restoreCell(s, 0, 0, prev)
    expect(s.cells.get(cellKey(0, 0))).toBe(prev)
    restoreCell(s, 0, 0, null)
    expect(s.cells.has(cellKey(0, 0))).toBe(false)
  })

  it('internStyle 去重', () => {
    const s = sheet()
    const a = { font: { bold: true }, numFmt: 'General' } as unknown as CellStyle
    const i1 = internStyle(s, a)
    const i2 = internStyle(s, { font: { bold: true }, numFmt: 'General' } as unknown as CellStyle)
    expect(i1).toBe(i2) // 深相等复用
  })
})

describe('mutations 维度(列宽/行高;E3.5)', () => {
  it('setColumnWidth/setRowHeight 写 Map + 保留 hidden + 下限保护', () => {
    const s = sheet()
    s.columns.set(0, { width: 50, hidden: true })
    setColumnWidth(s, 0, 100)
    expect(s.columns.get(0)).toMatchObject({ width: 100, hidden: true }) // hidden 保留
    setColumnWidth(s, 1, 2) // 低于下限
    expect(s.columns.get(1)!.width).toBe(8) // 夹到 8
    setRowHeight(s, 0, 30)
    expect(s.rows.get(0)).toMatchObject({ height: 30, hidden: false })
  })

  it('restoreDimension:写回信息 / null 删项(回落默认)', () => {
    const s = sheet()
    setColumnWidth(s, 0, 120)
    restoreDimension(s, 'col', 0, { width: 64, hidden: false })
    expect(s.columns.get(0)).toMatchObject({ width: 64 })
    restoreDimension(s, 'col', 0, null)
    expect(s.columns.has(0)).toBe(false) // null → 删项
    restoreDimension(s, 'row', 3, { height: 22, hidden: false })
    expect(s.rows.get(3)).toMatchObject({ height: 22 })
  })
})

describe('mutations 样式编辑(E5)', () => {
  it('mergeStyleOverride:font/fill 浅合并,其余覆盖', () => {
    const base = { font: { size: 11, color: '#000' }, fill: { type: 'none' }, hAlign: 'left', numFmt: 'General' } as unknown as CellStyle
    const merged = mergeStyleOverride(base, { font: { bold: true }, hAlign: 'center' })
    expect(merged.font).toMatchObject({ size: 11, color: '#000', bold: true }) // 浅合并保留旧字段
    expect(merged.hAlign).toBe('center') // 覆盖
  })

  it('applyStyleOverride:既有格改 styleId + intern 去重', () => {
    const s = sheet()
    s.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'x', styleId: 0 } as never)
    s.cells.set(cellKey(0, 1), { row: 0, col: 1, type: 'string', raw: 'y', styleId: 0 } as never)
    applyStyleOverride(s, 0, 0, { font: { bold: true } })
    const id0 = s.cells.get(cellKey(0, 0))!.styleId
    expect(id0).not.toBe(0) // 改了 styleId
    expect((s.styles[id0].font as { bold?: boolean }).bold).toBe(true)
    applyStyleOverride(s, 0, 1, { font: { bold: true } }) // 同基同补丁 → intern 复用
    expect(s.cells.get(cellKey(0, 1))!.styleId).toBe(id0)
  })

  it('applyStyleOverride:空格上色 → 新建 type=empty 格承载 styleId + 增长 dimension', () => {
    const s = sheet()
    applyStyleOverride(s, 5, 5, { fill: { type: 'solid', fgColor: '#ff0' } })
    const cell = s.cells.get(cellKey(5, 5))!
    expect(cell.type).toBe('empty')
    expect((s.styles[cell.styleId].fill as { fgColor?: string }).fgColor).toBe('#ff0')
    expect(s.dimension).toMatchObject({ rows: 6, cols: 6 })
  })
})
