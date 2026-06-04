import { describe, expect, it } from 'vitest'
import { makeSheet, makeStyle } from '../../__tests__/helpers'
import { cellKey } from '../types'
import type { CellModel, SheetModel } from '../types'
import {
  getCellStyle,
  getCellText,
  getCellValue,
  getRangeData,
  getSheetData,
  getWorkbookJSON,
  sheetToJSON,
} from '../data-access'

/** 构造一张带格式/日期/合并的表 */
function sample(): SheetModel {
  const styles = [
    makeStyle(), // 0: General
    makeStyle({ numFmt: '0.00%' }), // 1: 百分比
    makeStyle({ numFmt: 'yyyy"年"m"月"d"日"' }), // 2: 中文日期
  ]
  const cells = new Map<string, CellModel>()
  const put = (r: number, c: number, type: CellModel['type'], raw: CellModel['raw'], styleId = 0) =>
    cells.set(cellKey(r, c), { row: r, col: c, type, raw, styleId })

  // 表头
  put(0, 0, 'string', '产品')
  put(0, 1, 'string', '单价')
  put(0, 2, 'string', '占比')
  // D0 留空 → JSON 表头回退列字母 D
  // 数据
  put(1, 0, 'string', '鼠标')
  put(1, 1, 'number', 89)
  put(1, 2, 'number', 0.8734, 1)
  put(1, 3, 'date', new Date(Date.UTC(2026, 0, 15)), 2)
  put(2, 0, 'string', '键盘')
  put(2, 1, 'number', 399)
  put(2, 2, 'number', 0.15, 1)
  // 合并: A4:C4 锚点 A4='合计',覆盖格为空
  put(3, 0, 'string', '合计')

  return makeSheet({
    dimension: { rows: 4, cols: 4 },
    cells,
    styles,
    merges: [{ top: 3, left: 0, bottom: 3, right: 2 }],
  })
}

describe('逐格访问器', () => {
  const s = sample()
  it('getCellValue 取原始值', () => {
    expect(getCellValue(s, 1, 1)).toBe(89)
    expect(getCellValue(s, 1, 2)).toBe(0.8734)
    expect(getCellValue(s, 9, 9)).toBeNull() // 空格
  })
  it('getCellText 套数字/日期格式', () => {
    expect(getCellText(s, 1, 1)).toBe('89')
    expect(getCellText(s, 1, 2)).toBe('87.34%') // 0.8734 + 0.00%
    expect(getCellText(s, 1, 3)).toContain('2026') // 中文日期
    expect(getCellText(s, 1, 3)).toContain('年')
    expect(getCellText(s, 9, 9)).toBe('')
  })
  it('getCellStyle 解析 styleId', () => {
    expect(getCellStyle(s, 1, 2)?.numFmt).toBe('0.00%')
    expect(getCellStyle(s, 9, 9)).toBeUndefined()
  })
})

describe('整表 2D 数组', () => {
  const s = sample()
  it('稠密尺寸 + 默认显示文本', () => {
    const d = getSheetData(s)
    expect(d.length).toBe(4)
    expect(d[0].length).toBe(4)
    expect(d[1][1]).toBe('89')
    expect(d[1][2]).toBe('87.34%')
    expect(d[1][3]).toContain('年')
  })
  it('format:false → 原始值', () => {
    const d = getSheetData(s, { format: false })
    expect(d[1][1]).toBe(89)
    expect(d[1][2]).toBe(0.8734)
    expect(d[1][3]).toBeInstanceOf(Date)
  })
  it('合并: 锚点持值,覆盖格为空', () => {
    const d = getSheetData(s)
    expect(d[3][0]).toBe('合计')
    expect(d[3][1]).toBe('')
    expect(d[3][2]).toBe('')
  })
})

describe('区域 getRangeData', () => {
  it('子集(原始值)', () => {
    const s = sample()
    const d = getRangeData(s, { top: 1, left: 0, bottom: 2, right: 1 }, { format: false })
    expect(d).toEqual([
      ['鼠标', 89],
      ['键盘', 399],
    ])
  })
})

describe('sheetToJSON', () => {
  const s = sample()
  it('首行作 key(空表头回退列字母),默认显示文本', () => {
    const rows = sheetToJSON(s)
    expect(rows.length).toBe(3) // 鼠标/键盘/合计(合计行 A 非空,纳入)
    expect(rows[0]['产品']).toBe('鼠标')
    expect(rows[0]['占比']).toBe('87.34%')
    expect(Object.keys(rows[0])).toContain('D') // D0 空 → 列字母回退
  })
  it('format:false → 原始值', () => {
    const rows = sheetToJSON(s, { format: false })
    expect(rows[0]['单价']).toBe(89)
  })
  it('全空数据行被跳过', () => {
    const s2 = sample()
    s2.dimension = { rows: 6, cols: 4 } // 第 4、5 行全空
    expect(sheetToJSON(s2).length).toBe(3)
  })
})

describe('getWorkbookJSON', () => {
  it('可见表 → { 表名: 对象数组 }', () => {
    const s = sample()
    const wb = { sheets: [s], activeSheet: 0, themeColors: [], date1904: false }
    const out = getWorkbookJSON(wb)
    expect(Object.keys(out)).toEqual(['Sheet1'])
    expect(out['Sheet1'][0]['产品']).toBe('鼠标')
  })
})
