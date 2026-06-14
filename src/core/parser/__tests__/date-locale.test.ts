import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { parseWorkbook } from '../index'
import { cellKey } from '../../model/types'
import { formatValue } from '../../format/number-format'

// 回归:ExcelJS 把 OOXML 内置短日期(numFmtId 14)硬编码成美式 `mm-dd-yy`(→渲染 04-01-26),
// 我们在 adapter 重映射成中文 locale `yyyy/m/d`,跟 WPS/Excel 中文渲染对齐。
describe('内置短日期格式 locale 重映射(WPS 1:1)', () => {
  it('mm-dd-yy(ExcelJS 内置14)→ yyyy/m/d;日期渲染成 2026/4/1', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    const cell = ws.getCell('A1')
    cell.value = new Date(Date.UTC(2026, 3, 1)) // 2026-04-01
    cell.numFmt = 'mm-dd-yy' // = ExcelJS 对内置 14 的输出串
    const buf = await wb.xlsx.writeBuffer()

    const model = await parseWorkbook(buf as ArrayBuffer)
    const c = model.sheets[0].cells.get(cellKey(0, 0))!
    const st = model.sheets[0].styles[c.styleId]
    expect(st.numFmt).toBe('yyyy/m/d') // 重映射生效
    expect(formatValue(c.raw as number | Date, st.numFmt, model.date1904).text).toBe('2026/4/1')
  })

  it('空但带边框的格也入模型(结构格边框不丢)', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('B2').value = 'x'
    ws.getCell('C2').border = { top: { style: 'thin' } } // 空格,只有上边框
    const buf = await wb.xlsx.writeBuffer()
    const model = await parseWorkbook(buf as ArrayBuffer)
    const c2 = model.sheets[0].cells.get(cellKey(1, 2)) // C2
    expect(c2).toBeTruthy() // 空但带边框 → 入模型(以前 includeEmpty:false 会丢)
    expect(model.sheets[0].styles[c2!.styleId].borders.top?.style).toBe('thin')
    // 对照:真正空白格(无边框无填充)不入模型,不膨胀
    expect(model.sheets[0].cells.get(cellKey(5, 5))).toBeUndefined()
  })

  it('合并区锚点格的四边边框解析存活(合并边框存锚点)', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.mergeCells('A1:C1')
    ws.getCell('A1').value = '标题'
    ws.getCell('A1').border = { top: { style: 'thin' }, bottom: { style: 'medium' }, left: { style: 'thin' }, right: { style: 'thin' } }
    const buf = await wb.xlsx.writeBuffer()
    const model = await parseWorkbook(buf as ArrayBuffer)
    const s = model.sheets[0]
    const a1 = s.cells.get(cellKey(0, 0))!
    const b = s.styles[a1.styleId].borders
    expect(b.top?.style).toBe('thin')
    expect(b.bottom?.style).toBe('medium')
    expect(b.left?.style).toBe('thin')
    expect(b.right?.style).toBe('thin')
    expect(s.merges.some((m) => m.top === 0 && m.left === 0 && m.right === 2)).toBe(true)
  })

  it('列表型数据验证:内联选项被解析进 dataValidationLists', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').dataValidation = { type: 'list', allowBlank: true, formulae: ['"苹果,香蕉,橙子"'] }
    const buf = await wb.xlsx.writeBuffer()
    const model = await parseWorkbook(buf as ArrayBuffer)
    const lists = model.sheets[0].dataValidationLists
    expect(lists?.length).toBeGreaterThan(0)
    const hit = lists?.find((l) => l.range.top === 0 && l.range.left === 0)
    expect(hit?.options).toEqual(['苹果', '香蕉', '橙子'])
    expect(model.sheets[0].dataValidations.some((r) => r.top === 0 && r.left === 0)).toBe(true) // 箭头区域仍在
  })

  it('全类型数据验证:whole/decimal/date 规则连同 operator/formulae/messages 入 dataValidationRules', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').dataValidation = {
      type: 'whole', operator: 'between', allowBlank: false, formulae: [1, 100],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: '越界', error: '请输入 1-100',
      showInputMessage: true, promptTitle: '提示', prompt: '填 1 到 100',
    }
    ws.getCell('B1').dataValidation = { type: 'decimal', operator: 'greaterThan', allowBlank: true, formulae: [0] }
    const buf = await wb.xlsx.writeBuffer()
    const model = await parseWorkbook(buf as ArrayBuffer)
    const rules = model.sheets[0].dataValidationRules
    expect(rules?.length).toBeGreaterThanOrEqual(2)
    const whole = rules?.find((r) => r.type === 'whole' && r.range.top === 0 && r.range.left === 0)
    expect(whole?.operator).toBe('between')
    expect(whole?.formulae).toEqual([1, 100])
    expect(whole?.allowBlank).toBe(false)
    expect(whole?.error).toBe('请输入 1-100')
    expect(whole?.errorTitle).toBe('越界')
    expect(whole?.showInputMessage).toBe(true)
    expect(whole?.prompt).toBe('填 1 到 100')
    const dec = rules?.find((r) => r.type === 'decimal')
    expect(dec?.operator).toBe('greaterThan')
  })

  it('普通自定义/货币格式不受影响', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('S')
    ws.getCell('A1').value = 1234.5
    ws.getCell('A1').numFmt = '#,##0.00'
    const buf = await wb.xlsx.writeBuffer()
    const model = await parseWorkbook(buf as ArrayBuffer)
    const c = model.sheets[0].cells.get(cellKey(0, 0))!
    expect(model.sheets[0].styles[c.styleId].numFmt).toBe('#,##0.00')
  })
})
