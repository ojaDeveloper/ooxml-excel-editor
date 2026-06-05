import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { workbookToXlsxBlob } from '../xlsx-writer'
import type { CellStyle, SheetModel, WorkbookModel } from '../../model/types'
import { cellKey } from '../../model/types'

function style(over: Partial<CellStyle> = {}): CellStyle {
  return {
    font: { name: 'Arial', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#000000' },
    fill: { type: 'none' },
    borders: {},
    hAlign: 'general',
    vAlign: 'bottom',
    wrapText: false,
    shrinkToFit: false,
    textRotation: 0,
    indent: 0,
    numFmt: 'General',
    ...over,
  } as CellStyle
}

function workbook(): WorkbookModel {
  const styles = [style(), style({ font: { name: 'Arial', size: 11, bold: true, italic: false, underline: false, strike: false, color: '#FF0000' } })]
  const cells = new Map([
    [cellKey(0, 0), { row: 0, col: 0, type: 'number', raw: 42, styleId: 1 }],
    [cellKey(0, 1), { row: 0, col: 1, type: 'formula', raw: 84, formula: '=A1*2', styleId: 0 }],
    [cellKey(1, 0), { row: 1, col: 0, type: 'string', raw: 'hi', styleId: 0 }],
  ])
  const sheet = {
    name: 'S1',
    index: 0,
    state: 'visible',
    cells,
    styles,
    merges: [{ top: 2, left: 0, bottom: 2, right: 1 }],
    columns: new Map([[0, { width: 69, hidden: false }]]),
    rows: new Map([[0, { height: 30, hidden: false }]]),
    images: [],
    freeze: { frozenRows: 1, frozenCols: 1 },
    dimension: { rows: 3, cols: 2 },
  } as unknown as SheetModel
  return { sheets: [sheet], activeSheet: 0, themeColors: [], date1904: false } as unknown as WorkbookModel
}

describe('xlsx-writer 往返(从模型重建 → 重解析,值/样式/合并/几何存活)', () => {
  it('解析样例→写→重解析:值 + 加粗 + 公式 + 合并 + 列宽 + 冻结', async () => {
    const blob = await workbookToXlsxBlob(workbook())
    const buf = await blob.arrayBuffer()
    const wb2 = new ExcelJS.Workbook()
    await wb2.xlsx.load(buf)
    const ws = wb2.getWorksheet('S1')!

    expect(ws.getCell(1, 1).value).toBe(42) // 数值存活
    expect(ws.getCell(1, 1).font?.bold).toBe(true) // 加粗存活
    expect(ws.getCell(2, 1).value).toBe('hi') // 字符串存活
    const f = ws.getCell(1, 2).value as { formula: string; result: number }
    expect(f.formula).toBe('A1*2') // 公式文本(去 = )
    expect(f.result).toBe(84) // 缓存结果
    expect(ws.model.merges).toContain('A3:B3') // 合并存活
    expect(ws.getColumn(1).width).toBeGreaterThan(8) // 列宽(px→字符)合理
    expect(ws.views?.[0]?.state).toBe('frozen') // 冻结存活
  })

  it('编辑一格再导出:新值进入 xlsx', async () => {
    const wb = workbook()
    wb.sheets[0].cells.set(cellKey(1, 0), { row: 1, col: 0, type: 'string', raw: 'EDITED', styleId: 0 } as never)
    const wb2 = new ExcelJS.Workbook()
    await wb2.xlsx.load(await (await workbookToXlsxBlob(wb)).arrayBuffer())
    expect(wb2.getWorksheet('S1')!.getCell(2, 1).value).toBe('EDITED')
  })
})
