import { describe, it, expect } from 'vitest'
import { serializeSnapshot, encodeSnapshot, decodeSnapshot, reviveClipRaw, reattachImages, withoutImages } from '../clipboard-snapshot'
import { EditController, type EditControllerHost } from '../edit-controller'
import { makeSheet, makeStyle } from '../../__tests__/helpers'
import { cellKey } from '../../model/types'
import type { CellModel, SheetModel, WorkbookModel } from '../../model/types'

/** 造一张有"高保真内容"的源表:货币数字 + 红粗体 + 上边框 + 合并 + DISPIMG 格 + 浮动图 + 自定义行高列宽。 */
function richSource(): { sheet: SheetModel; wb: WorkbookModel } {
  const sheet = makeSheet({ name: 'Src' })
  const styled = makeStyle({
    font: { name: 'Calibri', size: 11, bold: true, italic: false, underline: false, strike: false, color: '#FF0000' },
    numFmt: '"¥"#,##0',
    borders: { top: { style: 'thin', color: '#000000' } },
  })
  sheet.styles = [makeStyle(), styled]
  const put = (cell: CellModel) => sheet.cells.set(cellKey(cell.row, cell.col), cell)
  put({ row: 0, col: 0, type: 'string', raw: '采购单号', styleId: 1 })
  put({ row: 0, col: 1, type: 'string', raw: '金额', styleId: 1 })
  put({ row: 1, col: 0, type: 'formula', raw: null, formula: '_xlfn.DISPIMG("ID_x",1)', dispImgId: 'ID_x', styleId: 0 })
  put({ row: 1, col: 1, type: 'number', raw: 237, styleId: 1 }) // 货币数字(关键:不能变成文本)
  sheet.merges = [{ top: 0, left: 0, bottom: 0, right: 1 }] // A1:B1 合并
  sheet.rows.set(1, { height: 88, hidden: false, customHeight: true })
  sheet.columns.set(1, { width: 120, hidden: false })
  sheet.images = [{ src: '', bytes: new Uint8Array([1, 2, 3, 4]), mime: 'image/png', from: { row: 1, col: 0, colOffEmu: 0, rowOffEmu: 0 }, extWidthEmu: 100, extHeightEmu: 100, editAs: 'oneCell' }]
  const wb: WorkbookModel = { sheets: [sheet], activeSheet: 0, themeColors: [], date1904: false, cellImages: new Map([['ID_x', { id: 'ID_x', bytes: new Uint8Array([9, 9, 9]), mime: 'image/png', src: '' }]]) }
  return { sheet, wb }
}

/** EditController host:目标表(空),全可编辑。 */
function targetController(): { sheet: SheetModel; wb: WorkbookModel; ec: EditController } {
  const sheet = makeSheet({ name: 'Tgt' })
  const wb: WorkbookModel = { sheets: [sheet], activeSheet: 0, themeColors: [], date1904: false }
  const host: EditControllerHost = {
    getSheet: () => sheet,
    getWorkbook: () => wb,
    getDate1904: () => false,
    isEditable: () => true,
    isEditingEnabled: () => true,
    getActiveSheetIndex: () => 0,
    isRecalcEnabled: () => false,
    getEngineFactory: () => null,
    onModelChange: () => {},
    rebuildOverlays: () => {},
    emit: () => {},
  }
  return { sheet, wb, ec: new EditController(host) }
}

describe('clipboard 1:1 保真快照', () => {
  it('serialize → encode → decode 往返:值/数字格式/边框/合并/图片/行高列宽 全保留', () => {
    const { sheet, wb } = richSource()
    const snap = decodeSnapshot(encodeSnapshot(serializeSnapshot(sheet, wb, { top: 0, left: 0, bottom: 1, right: 1 })))
    expect(snap).toBeTruthy()
    expect(snap!.rows).toBe(2)
    expect(snap!.cols).toBe(2)

    const num = snap!.cells.find((c) => c.r === 1 && c.c === 1)!
    expect(num.type).toBe('number')
    expect(num.raw).toBe(237) // 数字仍是数字,不是格式化文本
    expect(num.style.numFmt).toBe('"¥"#,##0') // 数字格式保留
    expect(num.style.font.bold).toBe(true)
    expect(num.style.font.color).toBe('#FF0000')
    expect(num.style.borders.top?.style).toBe('thin') // 边框保留(HTML 路径丢这个)

    const img = snap!.cells.find((c) => c.dispImgId === 'ID_x')!
    expect(img.formula).toContain('DISPIMG')
    expect(snap!.cellImages.find((ci) => ci.id === 'ID_x')).toBeTruthy() // DISPIMG 字节随快照带上

    expect(snap!.merges).toContainEqual({ top: 0, left: 0, bottom: 0, right: 1 })
    expect(snap!.rowHeights.find((d) => d.i === 1)).toMatchObject({ height: 88, custom: true })
    expect(snap!.colWidths.find((d) => d.i === 1)).toMatchObject({ width: 120 })
    expect(snap!.images).toHaveLength(1)
  })

  it('reviveClipRaw 还原 Date,其余原样', () => {
    const t = 1700000000000
    const d = reviveClipRaw({ __d: t })
    expect(d instanceof Date).toBe(true)
    expect((d as Date).getTime()).toBe(t)
    expect(reviveClipRaw(237)).toBe(237)
    expect(reviveClipRaw('x')).toBe('x')
    expect(reviveClipRaw(null)).toBe(null)
  })

  it('pasteSnapshot 覆盖式落到目标表 → 1:1(数字/格式/边框/合并/dispImgId/行高列宽/图片)', () => {
    const src = richSource()
    const snap = decodeSnapshot(encodeSnapshot(serializeSnapshot(src.sheet, src.wb, { top: 0, left: 0, bottom: 1, right: 1 })))!
    const tgt = targetController()
    const ok = tgt.ec.pasteSnapshot({ row: 5, col: 3 }, snap)
    expect(ok).toBe(true)

    // 数字格 → 目标 (6,4):仍是数字 237 + 货币格式 + 红粗 + 上边框
    const numCell = tgt.sheet.cells.get(cellKey(6, 4))!
    expect(numCell.type).toBe('number')
    expect(numCell.raw).toBe(237)
    const numStyle = tgt.sheet.styles[numCell.styleId]
    expect(numStyle.numFmt).toBe('"¥"#,##0')
    expect(numStyle.font.bold).toBe(true)
    expect(numStyle.borders.top?.style).toBe('thin')

    // DISPIMG 格 → 目标 (6,3):dispImgId 保留 + 目标 cellImages 已登记字节
    const imgCell = tgt.sheet.cells.get(cellKey(6, 3))!
    expect(imgCell.dispImgId).toBe('ID_x')
    expect(tgt.wb.cellImages?.get('ID_x')?.bytes).toBeTruthy()

    // 合并 → A1:B1 平移到 (5,3)-(5,4)
    expect(tgt.sheet.merges).toContainEqual({ top: 5, left: 3, bottom: 5, right: 4 })
    // 行高列宽平移
    expect(tgt.sheet.rows.get(6)).toMatchObject({ height: 88, customHeight: true })
    expect(tgt.sheet.columns.get(4)).toMatchObject({ width: 120 })
    // 浮动图平移到 from (6,3)
    expect(tgt.sheet.images.some((im) => im.from.row === 6 && im.from.col === 3)).toBe(true)

    // 单次撤销 → 目标全清回去
    tgt.ec.undo()
    expect(tgt.sheet.cells.get(cellKey(6, 4))).toBeUndefined()
    expect(tgt.sheet.merges).toHaveLength(0)
    expect(tgt.sheet.images).toHaveLength(0)
  })

  it('瘦身传输:withImageBytes=false 去图片字节,reattachImages 从 <img> 回填 → pasteSnapshot 仍 1:1', () => {
    const { sheet, wb } = richSource()
    const range = { top: 0, left: 0, bottom: 1, right: 1 }
    const full = serializeSnapshot(sheet, wb, range)
    const lite = serializeSnapshot(sheet, wb, range, { withImageBytes: false })
    expect(lite.cellImages[0].b64).toBe('') // 瘦身:快照里不带图片字节(避免双重 base64)
    expect(lite.images[0].b64).toBe('')

    // 模拟可见 <img data-clip-img> 携带的字节(key: DISPIMG=c:id,浮动=f:序号)
    const imgB64 = new Map<string, string>([
      ['c:ID_x', full.cellImages.find((c) => c.id === 'ID_x')!.b64],
      ['f:0', full.images[0].b64],
    ])
    const reattached = reattachImages(decodeSnapshot(encodeSnapshot(lite)), imgB64)!
    expect(reattached.cellImages.find((c) => c.id === 'ID_x')!.b64).toBe(imgB64.get('c:ID_x'))
    expect(reattached.images[0].b64).toBe(imgB64.get('f:0'))

    const tgt = targetController()
    tgt.ec.pasteSnapshot({ row: 0, col: 0 }, reattached)
    expect(tgt.wb.cellImages?.get('ID_x')?.bytes).toBeTruthy() // 图片字节回来了
    expect(tgt.sheet.images).toHaveLength(1)
  })

  it('withoutImages 降级:去掉所有图片 + DISPIMG 格中性化为空格(保其余 1:1)', () => {
    const { sheet, wb } = richSource()
    const snap = withoutImages(serializeSnapshot(sheet, wb, { top: 0, left: 0, bottom: 1, right: 1 }))
    expect(snap.cellImages).toHaveLength(0)
    expect(snap.images).toHaveLength(0)
    const wasDisp = snap.cells.find((c) => c.r === 1 && c.c === 0)!
    expect(wasDisp.dispImgId).toBeUndefined()
    expect(wasDisp.type).toBe('empty')
    // 非图片格仍在(数字格 1:1 保留)
    expect(snap.cells.find((c) => c.r === 1 && c.c === 1)?.raw).toBe(237)
  })

  it('decodeSnapshot 对脏数据/非本组件 HTML 返 null(回退外部解析)', () => {
    expect(decodeSnapshot('not-base64!!')).toBeNull()
    expect(decodeSnapshot('')).toBeNull()
    expect(decodeSnapshot(null)).toBeNull()
  })
})
