import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { applyCommand } from '../commands'
import { parseWorkbook } from '../../parser/index'
import { workbookToXlsxBlob } from '../../export/xlsx-writer'
import { cellKey, type SheetModel } from '../../model/types'

function emptySheet(): SheetModel {
  return {
    name: 'S', index: 0, state: 'visible', dimension: { rows: 2, cols: 2 },
    cells: new Map(), styles: [], merges: [], columns: new Map(), rows: new Map(),
    defaultColWidth: 64, defaultRowHeight: 20, freeze: { frozenRows: 0, frozenCols: 0 },
    conditional: [], dataValidations: [], images: [], charts: [], shapes: [], sparklines: [], pivotTables: [],
  } as unknown as SheetModel
}

describe('批注编辑(1.11.0):命令可撤销', () => {
  it('set-comment 给空格挂批注;逆命令还原(删格)', () => {
    const sheet = emptySheet()
    const { inverse } = applyCommand(sheet, { kind: 'set-comment', row: 1, col: 1, comment: '复核一下' })
    expect(sheet.cells.get(cellKey(1, 1))?.comment).toBe('复核一下')
    applyCommand(sheet, inverse) // undo
    expect(sheet.cells.get(cellKey(1, 1))).toBeUndefined() // 空格批注删掉后整格清掉
  })

  it('空批注清除已有批注', () => {
    const sheet = emptySheet()
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'x', styleId: 0, comment: '旧' } as never)
    applyCommand(sheet, { kind: 'set-comment', row: 0, col: 0, comment: '' })
    expect(sheet.cells.get(cellKey(0, 0))?.comment).toBeUndefined()
    expect(sheet.cells.get(cellKey(0, 0))?.raw).toBe('x') // 有值的格保留
  })
})

describe('批注编辑(1.11.0):导出回写往返', () => {
  it('rebuild:模型批注导出后再解析仍在', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').value = 'hi'
    const buf = await wb.xlsx.writeBuffer()
    const model = await parseWorkbook(buf as ArrayBuffer)
    // app 内加批注
    model.sheets[0].cells.get(cellKey(0, 0))!.comment = '导出测试批注'
    const blob = await workbookToXlsxBlob(model, {})
    const re = await parseWorkbook(await blob.arrayBuffer())
    expect(re.sheets[0].cells.get(cellKey(0, 0))?.comment).toContain('导出测试批注')
  })
})
