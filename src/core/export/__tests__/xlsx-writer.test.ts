import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import ExcelJS from 'exceljs'
import { workbookToXlsxBlob } from '../xlsx-writer'
import { parseWorkbook } from '../../parser/index'
import type { CellStyle, SheetModel, WorkbookModel } from '../../model/types'
import { cellKey } from '../../model/types'

function loadSample(): ArrayBuffer {
  const buf = readFileSync(join(__dirname, '..', '..', '..', '..', 'public', 'sample.xlsx'))
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)
}

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

  it('富文本往返:每段字体(颜色/粗体)不丢', async () => {
    const wb = workbook()
    wb.sheets[0].cells.set(cellKey(1, 1), {
      row: 1, col: 1, type: 'richtext', raw: 'AB',
      rich: [
        { text: 'A', font: { color: '#FF0000', bold: true } },
        { text: 'B', font: { color: '#0000FF' } },
      ],
      styleId: 0,
    } as never)
    const buf = await (await workbookToXlsxBlob(wb)).arrayBuffer()
    const re = await parseWorkbook(buf)
    const cell = re.sheets[0].cells.get(cellKey(1, 1))!
    expect(cell.type).toBe('richtext')
    expect(cell.rich?.length).toBe(2)
    expect(String(cell.rich?.[0].font?.color).toUpperCase()).toBe('#FF0000') // 红(以前导出丢)
    expect(cell.rich?.[0].font?.bold).toBe(true)
    expect(String(cell.rich?.[1].font?.color).toUpperCase()).toBe('#0000FF') // 蓝
  })

  it('编辑一格再导出:新值进入 xlsx', async () => {
    const wb = workbook()
    wb.sheets[0].cells.set(cellKey(1, 0), { row: 1, col: 0, type: 'string', raw: 'EDITED', styleId: 0 } as never)
    const wb2 = new ExcelJS.Workbook()
    await wb2.xlsx.load(await (await workbookToXlsxBlob(wb)).arrayBuffer())
    expect(wb2.getWorksheet('S1')!.getCell(2, 1).value).toBe('EDITED')
  })

  it('图片导出:twoCell 锚出 br(随格缩放),oneCell 锚出 ext(F2)', async () => {
    const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
    const wb = workbook()
    const s = wb.sheets[0] as any
    s.defaultColWidth = 64
    s.defaultRowHeight = 20
    s.images = [
      { src: PNG, from: { col: 1, colOffEmu: 0, row: 1, rowOffEmu: 0 }, to: { col: 3, colOffEmu: 0, row: 2, rowOffEmu: 0 } }, // twoCell
      { src: PNG, from: { col: 0, colOffEmu: 95250, row: 0, rowOffEmu: 0 }, extWidthEmu: 952500, extHeightEmu: 952500 }, // oneCell + 偏移
    ]
    const wb2 = new ExcelJS.Workbook()
    await wb2.xlsx.load(await (await workbookToXlsxBlob(wb)).arrayBuffer())
    const imgs = wb2.getWorksheet('S1')!.getImages()
    expect(imgs).toHaveLength(2)
    const twoCell = imgs.find((i) => i.range.br)!
    const oneCell = imgs.find((i) => !i.range.br)!
    expect(twoCell.range.br!.nativeCol).toBe(3) // br 到第 4 列
    expect(oneCell.range.tl.nativeColOff).toBeGreaterThan(0) // 子格 EMU 偏移保真(非 0)
  })
})

describe('xlsx-writer overlay 高保真(重载原件叠加;F3)', () => {
  it('overlay 与 rebuild 都保留原件条件格式(1.9.0 起 rebuild 也按模型回写);两者都反映编辑值', async () => {
    const src = loadSample()
    const model = await parseWorkbook(src) // 样例首表有 2 条条件格式
    model.sheets[0].cells.set(cellKey(2, 1), { row: 2, col: 1, type: 'number', raw: 88888, styleId: 0 } as never)

    // overlay:CF 保留 + 编辑值生效
    const ov = new ExcelJS.Workbook()
    await ov.xlsx.load(await (await workbookToXlsxBlob(model, { fidelity: 'overlay', sourceBuffer: src })).arrayBuffer())
    const ovCF = (ov.worksheets[0] as unknown as { conditionalFormattings?: unknown[] }).conditionalFormattings ?? []
    expect(ovCF.length).toBeGreaterThan(0) // 条件格式存活
    expect(ov.worksheets[0].getCell(3, 2).value).toBe(88888) // 编辑值叠加生效

    // rebuild:1.9.0 起也回写条件格式(parsed 规则用 raw 原样写),编辑值仍在
    const rb = new ExcelJS.Workbook()
    await rb.xlsx.load(await (await workbookToXlsxBlob(model, {})).arrayBuffer())
    const rbCF = (rb.worksheets[0] as unknown as { conditionalFormattings?: unknown[] }).conditionalFormattings ?? []
    expect(rbCF.length).toBeGreaterThan(0) // rebuild 不再丢条件格式
    expect(rb.worksheets[0].getCell(3, 2).value).toBe(88888)
  })

  it('overlay 缺 sourceBuffer → 回退 rebuild(不报错)', async () => {
    const model = await parseWorkbook(loadSample())
    const blob = await workbookToXlsxBlob(model, { fidelity: 'overlay' }) // 无 sourceBuffer
    expect(blob.size).toBeGreaterThan(1000)
  })
})
