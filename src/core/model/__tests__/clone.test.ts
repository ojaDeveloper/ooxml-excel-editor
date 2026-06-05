import { describe, it, expect } from 'vitest'
import { cloneWorkbook, restoreWorkbookInto } from '../clone'
import { setCellValue, setColumnWidth } from '../mutations'
import type { SheetModel, WorkbookModel } from '../types'
import { cellKey } from '../types'

function workbook(): WorkbookModel {
  const sheet = {
    name: 'S1',
    index: 0,
    state: 'visible',
    cells: new Map([[cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'orig', styleId: 0 }]]),
    styles: [],
    merges: [],
    columns: new Map([[0, { width: 64, hidden: false }]]),
    rows: new Map(),
    images: [],
    dimension: { rows: 1, cols: 1 },
    defaultColWidth: 64,
    defaultRowHeight: 20,
  } as unknown as SheetModel
  return { sheets: [sheet], activeSheet: 0, themeColors: [], date1904: false } as unknown as WorkbookModel
}

describe('clone — cloneWorkbook + restoreWorkbookInto(脏状态还原;E3.5)', () => {
  it('cloneWorkbook 深克隆(改克隆不影响原件)', () => {
    const wb = workbook()
    const snap = cloneWorkbook(wb)
    setCellValue(snap.sheets[0], 0, 0, 'changed')
    expect(wb.sheets[0].cells.get(cellKey(0, 0))).toMatchObject({ raw: 'orig' }) // 原件不动
  })

  it('restoreWorkbookInto:还原值 + 列宽,且保留 sheet 对象身份', () => {
    const wb = workbook()
    const baseline = cloneWorkbook(wb)
    const liveSheetRef = wb.sheets[0] // 还原后须仍是同一对象(壳/渲染器持有此引用)

    // 改:编辑一格 + 改列宽 + 加一格
    setCellValue(wb.sheets[0], 0, 0, 'edited')
    setColumnWidth(wb.sheets[0], 0, 200)
    setCellValue(wb.sheets[0], 5, 5, 'extra')

    restoreWorkbookInto(wb, baseline)
    expect(wb.sheets[0]).toBe(liveSheetRef) // 对象身份不变
    expect(wb.sheets[0].cells.get(cellKey(0, 0))).toMatchObject({ raw: 'orig' }) // 值还原
    expect(wb.sheets[0].columns.get(0)).toMatchObject({ width: 64 }) // 列宽还原
    expect(wb.sheets[0].cells.has(cellKey(5, 5))).toBe(false) // 新增格被清掉
  })

  it('baseline 可重复还原(还原后再改再还原仍回原件)', () => {
    const wb = workbook()
    const baseline = cloneWorkbook(wb)
    setColumnWidth(wb.sheets[0], 0, 300)
    restoreWorkbookInto(wb, baseline)
    setColumnWidth(wb.sheets[0], 0, 999)
    restoreWorkbookInto(wb, baseline) // 第二次还原
    expect(wb.sheets[0].columns.get(0)).toMatchObject({ width: 64 })
  })
})
