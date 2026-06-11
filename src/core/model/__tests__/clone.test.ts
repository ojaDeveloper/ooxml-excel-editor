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

describe('clone — 轻量快照:重负载共享、可变部分克隆(性能)', () => {
  it('图片字节/图表 按引用共享,锚点与格 克隆(不深拷重负载)', () => {
    const bytes = new Uint8Array([1, 2, 3, 4])
    const chart = { type: 'bar', series: [] }
    const wb = workbook()
    const s = wb.sheets[0] as any
    s.images = [{ src: 'x', bytes, mime: 'image/png', from: { col: 0, colOffEmu: 0, row: 0, rowOffEmu: 0 } }]
    s.charts = [chart]
    const c = cloneWorkbook(wb)
    const cs = c.sheets[0] as any
    expect(cs.images[0].bytes).toBe(bytes) // 字节共享(不深拷,省内存)
    expect(cs.images[0]).not.toBe(s.images[0]) // 锚点对象克隆(结构编辑挪锚点不污染快照)
    expect(cs.charts).toBe(s.charts) // 图表整体共享(编辑期间不可变)
    expect(cs.cells.get(cellKey(0, 0))).not.toBe(s.cells.get(cellKey(0, 0))) // 格克隆
    expect(cs.styles).not.toBe(s.styles) // styles 新数组(防 internStyle 追加污染)
  })

  it('pivotTables 元数据克隆,布局变更不污染快照', () => {
    const wb = workbook()
    wb.sheets[0].pivotTables = [{
      name: 'PivotTable1',
      range: { top: 0, left: 0, bottom: 2, right: 2 },
      fields: ['城市', '金额'],
      buttons: [{ row: 0, col: 0, label: '城市', kind: 'row' }],
      source: { sheetIndex: 0, range: { top: 0, left: 0, bottom: 10, right: 2 } },
      layout: { filters: [], columns: [], rows: [0], values: [{ field: 1, summary: 'sum' }] },
    }]
    const snap = cloneWorkbook(wb)
    wb.sheets[0].pivotTables[0].layout!.values[0].summary = 'avg'
    expect(snap.sheets[0].pivotTables[0].layout!.values[0].summary).toBe('sum')
  })
})

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

  it('restoreWorkbookInto:还原 sheet 数量,撤销新增工作表', () => {
    const wb = workbook()
    const baseline = cloneWorkbook(wb)
    wb.sheets.push({ ...wb.sheets[0], name: 'PivotTable', index: 1, cells: new Map() })
    wb.activeSheet = 1

    restoreWorkbookInto(wb, baseline)
    expect(wb.sheets).toHaveLength(1)
    expect(wb.activeSheet).toBe(0)
    expect(wb.sheets[0].name).toBe('S1')
  })
})
