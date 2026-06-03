import { describe, it, expect } from 'vitest'
import { autoFitRowHeights } from '../autofit'
import type { CellStyle, SheetModel, WorkbookModel } from '../../model/types'
import { cellKey } from '../../model/types'

// 模拟测量上下文: 每个字符宽 7px
function mockCtx(): CanvasRenderingContext2D {
  return {
    font: '',
    save() {},
    restore() {},
    measureText(t: string) {
      return { width: t.length * 7 } as TextMetrics
    },
  } as unknown as CanvasRenderingContext2D
}

const baseStyle = (over: Partial<CellStyle> = {}): CellStyle => ({
  font: { name: 'Calibri', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#000' },
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
})

function makeSheet(cell: { text: string; wrap: boolean }, colWidth: number): SheetModel {
  const styles = [baseStyle({ wrapText: cell.wrap })]
  const cells = new Map()
  cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: cell.text, styleId: 0 })
  return {
    name: 'S', index: 0, state: 'visible',
    dimension: { rows: 1, cols: 1 },
    cells, styles, merges: [],
    columns: new Map([[0, { width: colWidth, hidden: false }]]),
    rows: new Map(),
    defaultColWidth: 64, defaultRowHeight: 20,
    freeze: { frozenRows: 0, frozenCols: 0 },
    conditional: [], dataValidations: [], images: [], charts: [], shapes: [], sparklines: [], showGridLines: true,
  }
}

const wb: WorkbookModel = { sheets: [], activeSheet: 0, themeColors: [], date1904: false }

describe('autoFitRowHeights', () => {
  it('换行长文本 → 行被撑高', () => {
    const sheet = makeSheet({ text: 'A'.repeat(40), wrap: true }, 60)
    autoFitRowHeights(sheet, wb, mockCtx())
    const h = sheet.rows.get(0)?.height ?? 20
    expect(h).toBeGreaterThan(40) // 6 行左右，远高于默认 20
  })

  it('单行不换行 → 行高不变', () => {
    const sheet = makeSheet({ text: '短文本', wrap: false }, 200)
    autoFitRowHeights(sheet, wb, mockCtx())
    expect(sheet.rows.has(0)).toBe(false) // 未写入 → 沿用默认
  })

  it('显式 \\n 多行 → 撑高', () => {
    const sheet = makeSheet({ text: '第一行\n第二行\n第三行', wrap: false }, 200)
    autoFitRowHeights(sheet, wb, mockCtx())
    const h = sheet.rows.get(0)?.height ?? 20
    expect(h).toBeGreaterThan(40)
  })

  it('合并单元格被排除(不撑高)', () => {
    const sheet = makeSheet({ text: 'A'.repeat(40), wrap: true }, 60)
    sheet.merges = [{ top: 0, left: 0, bottom: 0, right: 0 }]
    autoFitRowHeights(sheet, wb, mockCtx())
    expect(sheet.rows.has(0)).toBe(false)
  })

  it('只扩不缩: 已有更大行高时保留', () => {
    const sheet = makeSheet({ text: 'A'.repeat(10), wrap: true }, 60)
    sheet.rows.set(0, { height: 200, hidden: false })
    autoFitRowHeights(sheet, wb, mockCtx())
    expect(sheet.rows.get(0)?.height).toBe(200)
  })
})
