/**
 * 模板样式 overlay 单测 —— 验证 applyStyleTemplate 的核心契约:
 *   1) 数据 raw 全部保留在自然位置 (A1 起)
 *   2) 模板 raw 全部丢弃 (装饰文字 / 占位符 / 表头都不带过来)
 *   3) styleId 从模板同位置取
 *   4) merges / 列宽 / 行高 / freeze / showGridLines / themeColors 全部从模板拷贝
 *   5) 模板的 images / charts / conditional / dataValidations 全部 NOT 带过来 (避免幽灵规则)
 *   6) 数据 sheet 名 / date1904 / cellImages 透传
 */
import { describe, it, expect } from 'vitest'
import { applyStyleTemplate } from '../style-overlay'
import { jsonToWorkbook } from '../../loader-json'
import { cellKey } from '../../model/types'
import type { CellStyle, SheetModel, WorkbookModel } from '../../model/types'

function makeStyle(overrides: Partial<CellStyle> = {}): CellStyle {
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
    ...overrides,
  }
}

/** 构造一个带 styling 的"发票模板": A1 大标题 / A4 表头 / A5:C9 边框数据区 / A11 合计 */
function makeInvoiceTemplate(): WorkbookModel {
  const sheet: SheetModel = {
    name: '发票模板',
    index: 0,
    state: 'visible',
    dimension: { rows: 12, cols: 3 },
    cells: new Map(),
    styles: [
      makeStyle(),                                                                  // 0: 默认
      makeStyle({ font: { ...makeStyle().font, bold: true, size: 18 } }),           // 1: 大标题
      makeStyle({ font: { ...makeStyle().font, bold: true }, fill: { type: 'solid', fgColor: '#FFF8DC' } }), // 2: 表头
      makeStyle({ borders: { top: { style: 'thin', color: '#000' }, bottom: { style: 'thin', color: '#000' } } }), // 3: 数据格
      makeStyle({ font: { ...makeStyle().font, bold: true } }),                     // 4: 合计行
    ],
    merges: [{ top: 0, left: 0, bottom: 0, right: 2 }],   // A1:C1 合并 (标题)
    columns: new Map([
      [0, { width: 120, hidden: false }],
      [1, { width: 80, hidden: false }],
      [2, { width: 100, hidden: false }],
    ]),
    rows: new Map([[0, { height: 32, hidden: false }]]),
    defaultColWidth: 64,
    defaultRowHeight: 20,
    freeze: { frozenRows: 4, frozenCols: 0 },     // 冻结表头区
    conditional: [{ ranges: [{ top: 4, left: 2, bottom: 8, right: 2 }], priority: 1, type: 'cellIs', operator: 'greaterThan', formulae: ['100'] }],
    dataValidations: [{ top: 4, left: 0, bottom: 8, right: 2 }],
    images: [],
    charts: [],
    shapes: [],
    sparklines: [],
    showGridLines: false,
  }
  // 装饰文字 (会被丢弃, 仅用于验证清空)
  sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: '订单结算单', styleId: 1 })
  sheet.cells.set(cellKey(1, 0), { row: 1, col: 0, type: 'string', raw: '客户:', styleId: 0 })
  sheet.cells.set(cellKey(1, 1), { row: 1, col: 1, type: 'string', raw: '{{customer}}', styleId: 0 })
  sheet.cells.set(cellKey(3, 0), { row: 3, col: 0, type: 'string', raw: '商品', styleId: 2 })
  sheet.cells.set(cellKey(3, 1), { row: 3, col: 1, type: 'string', raw: '单价', styleId: 2 })
  sheet.cells.set(cellKey(3, 2), { row: 3, col: 2, type: 'string', raw: '数量', styleId: 2 })
  // 模板预留的数据格 (有边框样式)
  for (let r = 4; r < 9; r++) {
    for (let c = 0; c < 3; c++) {
      sheet.cells.set(cellKey(r, c), { row: r, col: c, type: 'empty', raw: null, styleId: 3 })
    }
  }
  sheet.cells.set(cellKey(10, 0), { row: 10, col: 0, type: 'string', raw: '合计:', styleId: 4 })
  sheet.cells.set(cellKey(10, 2), { row: 10, col: 2, type: 'string', raw: '{{total}}', styleId: 4 })

  return {
    sheets: [sheet],
    activeSheet: 0,
    themeColors: Array(17).fill('#444444'),
    date1904: false,
  }
}

describe('applyStyleTemplate — 核心契约', () => {
  it('数据保留在自然位置 A1 起,模板装饰文字全部丢弃', () => {
    const data = jsonToWorkbook([
      { name: '笔记本电脑', price: 5999, qty: 2 },
      { name: '机械键盘', price: 399, qty: 1 },
    ])
    const tpl = makeInvoiceTemplate()
    const out = applyStyleTemplate(data, tpl)
    const sheet = out.sheets[0]

    // 数据在 A1 起 (jsonToWorkbook 自动加表头行)
    expect(sheet.cells.get(cellKey(0, 0))?.raw).toBe('name')
    expect(sheet.cells.get(cellKey(0, 1))?.raw).toBe('price')
    expect(sheet.cells.get(cellKey(1, 0))?.raw).toBe('笔记本电脑')
    expect(sheet.cells.get(cellKey(2, 0))?.raw).toBe('机械键盘')

    // 模板装饰文字一个都不见
    const allRaws = [...sheet.cells.values()].map((c) => c.raw)
    expect(allRaws).not.toContain('订单结算单')
    expect(allRaws).not.toContain('客户:')
    expect(allRaws).not.toContain('{{customer}}')
    expect(allRaws).not.toContain('商品')
    expect(allRaws).not.toContain('合计:')
    expect(allRaws).not.toContain('{{total}}')

    // 模板预留的空白边框格 (A5:C9 4-8 行) 也不见 —— 它们只是模板的空格,数据没填的位置不该有"空 raw 格"
    expect(sheet.cells.get(cellKey(4, 0))).toBeUndefined()
    expect(sheet.cells.get(cellKey(8, 2))).toBeUndefined()
  })

  it('styleId 从模板同位置取;数据格落在模板"装饰区"会继承装饰区的 styleId', () => {
    const data = jsonToWorkbook([['标题占用 A1']])    // A1 单元格
    const tpl = makeInvoiceTemplate()                  // 模板 A1 是 styleId=1 (大标题)
    const out = applyStyleTemplate(data, tpl)
    const sheet = out.sheets[0]

    // 数据 A1 继承模板 A1 的 styleId=1
    expect(sheet.cells.get(cellKey(0, 0))?.styleId).toBe(1)
    // styles 池来自模板
    expect(sheet.styles).toHaveLength(5)
    expect(sheet.styles[1].font.size).toBe(18)
  })

  it('数据格落在模板没有定义的位置 (超出模板范围) → styleId=0 默认', () => {
    const data = jsonToWorkbook([['x', 'y', 'z', '超出模板的第四列']])
    const tpl = makeInvoiceTemplate()
    const out = applyStyleTemplate(data, tpl)
    const sheet = out.sheets[0]
    // 模板没有 D 列 (col=3) 的定义,数据落上去 → styleId=0
    expect(sheet.cells.get(cellKey(0, 3))?.styleId).toBe(0)
    expect(sheet.cells.get(cellKey(0, 3))?.raw).toBe('超出模板的第四列')
  })

  it('merges / 列宽 / 行高 / freeze / showGridLines / themeColors 全部从模板拷贝', () => {
    const data = jsonToWorkbook([['x']])
    const tpl = makeInvoiceTemplate()
    const out = applyStyleTemplate(data, tpl)
    const sheet = out.sheets[0]

    expect(sheet.merges).toEqual([{ top: 0, left: 0, bottom: 0, right: 2 }])
    expect(sheet.columns.get(0)?.width).toBe(120)
    expect(sheet.columns.get(2)?.width).toBe(100)
    expect(sheet.rows.get(0)?.height).toBe(32)
    expect(sheet.freeze).toEqual({ frozenRows: 4, frozenCols: 0 })
    expect(sheet.showGridLines).toBe(false)
    expect(out.themeColors[0]).toBe('#444444')
  })

  it('模板的 images / charts / shapes / conditional / dataValidations 全部 NOT 带过来', () => {
    const data = jsonToWorkbook([['x']])
    const tpl = makeInvoiceTemplate()
    const out = applyStyleTemplate(data, tpl)
    const sheet = out.sheets[0]

    expect(sheet.images).toEqual([])
    expect(sheet.charts).toEqual([])
    expect(sheet.shapes).toEqual([])
    expect(sheet.conditional).toEqual([])
    expect(sheet.dataValidations).toEqual([])
  })

  it('数据的 sheet 名 / date1904 / cellImages 透传 (而不是用模板的)', () => {
    const data = jsonToWorkbook({ sheets: [{ name: '我的数据表', rows: [['x']] }] })
    data.date1904 = true
    data.cellImages = new Map([['id-1', { id: 'id-1', src: 'blob:x' }]])

    const tpl = makeInvoiceTemplate()
    const out = applyStyleTemplate(data, tpl)
    expect(out.sheets[0].name).toBe('我的数据表')  // 数据 sheet 名
    expect(out.date1904).toBe(true)                 // 数据 date1904
    expect(out.cellImages?.size).toBe(1)            // 数据 cellImages
  })

  it('数据 dimension 跟模板列宽行高声明取大 —— 模板设了 20 列宽但数据只 3 列,虚拟范围仍按模板撑出', () => {
    const data = jsonToWorkbook([['a', 'b', 'c']])  // 3 列
    const tpl = makeInvoiceTemplate()                 // columns 显式定义到 col=2 (3 列)
    // 额外加一列定义
    tpl.sheets[0].columns.set(10, { width: 80, hidden: false })
    const out = applyStyleTemplate(data, tpl)
    expect(out.sheets[0].dimension.cols).toBeGreaterThanOrEqual(11)
  })

  it('入参不被原地修改 (data + template 都保持纯净)', () => {
    const data = jsonToWorkbook([['x']])
    const tpl = makeInvoiceTemplate()
    const dataA1 = data.sheets[0].cells.get(cellKey(0, 0))?.raw
    const tplA1 = tpl.sheets[0].cells.get(cellKey(0, 0))?.raw

    applyStyleTemplate(data, tpl)

    expect(data.sheets[0].cells.get(cellKey(0, 0))?.raw).toBe(dataA1)
    expect(tpl.sheets[0].cells.get(cellKey(0, 0))?.raw).toBe(tplA1)
  })

  it('空数据 + 模板 → 干净的样式骨架 (无任何 raw,但 merges / 列宽 / freeze 全在)', () => {
    const data = jsonToWorkbook([])
    const tpl = makeInvoiceTemplate()
    const out = applyStyleTemplate(data, tpl)
    const sheet = out.sheets[0]
    expect(sheet.cells.size).toBe(0)
    expect(sheet.merges.length).toBe(1)
    expect(sheet.columns.get(0)?.width).toBe(120)
    expect(sheet.freeze.frozenRows).toBe(4)
  })
})
