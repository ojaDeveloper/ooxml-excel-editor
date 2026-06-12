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
