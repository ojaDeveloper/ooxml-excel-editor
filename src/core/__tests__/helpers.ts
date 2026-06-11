/** 测试辅助: 构造最小可用的 SheetModel / CellStyle / CellModel。 */
import type { CellModel, CellStyle, SheetModel } from '../model/types'
import { cellKey } from '../model/types'

export function makeStyle(over: Partial<CellStyle> = {}): CellStyle {
  return {
    font: { name: 'Calibri', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#000000' },
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
  }
}

export function makeSheet(over: Partial<SheetModel> = {}): SheetModel {
  const base: SheetModel = {
    name: 'Sheet1',
    index: 0,
    state: 'visible',
    dimension: { rows: 10, cols: 5 },
    cells: new Map(),
    styles: [makeStyle()],
    merges: [],
    columns: new Map(),
    rows: new Map(),
    defaultColWidth: 64,
    defaultRowHeight: 20,
    freeze: { frozenRows: 0, frozenCols: 0 },
    conditional: [],
    dataValidations: [],
    images: [],
    charts: [],
    shapes: [],
    sparklines: [],
    pivotTables: [],
    showGridLines: true,
  }
  return { ...base, ...over }
}

/** 往 sheet 填数值单元格(styleId=0) */
export function putNumbers(sheet: SheetModel, cells: { row: number; col: number; v: number }[]): void {
  for (const c of cells) {
    const m: CellModel = { row: c.row, col: c.col, type: 'number', raw: c.v, styleId: 0 }
    sheet.cells.set(cellKey(c.row, c.col), m)
  }
}
