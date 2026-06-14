import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { parseWorkbook } from '../../parser/index'
import { workbookToXlsxBlob } from '../../export/xlsx-writer'
import { applyCommand } from '../commands'
import type { ConditionalRule, SheetModel } from '../../model/types'

/** 造一个带多种条件格式的 .xlsx buffer。 */
async function fixtureBuffer(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('S')
  for (let i = 1; i <= 6; i++) ws.getCell('A' + i).value = i * 10
  ws.addConditionalFormatting({ ref: 'A1:A6', rules: [{ type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }], color: [{ argb: 'FFFF0000' }, { argb: 'FFFFFF00' }, { argb: 'FF00FF00' }], priority: 1 } as any] })
  ws.addConditionalFormatting({ ref: 'A1:A6', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['30'], priority: 2, style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } } } } as any] })
  ws.addConditionalFormatting({ ref: 'A1:A6', rules: [{ type: 'iconSet', iconSet: '3TrafficLights1', reverse: true, cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }], priority: 3 } as any] })
  ws.addConditionalFormatting({ ref: 'A1:A6', rules: [{ type: 'top10', rank: 3, percent: false, bottom: true, priority: 4, style: { font: { bold: true } } } as any] })
  const buf = await wb.xlsx.writeBuffer()
  return buf as ArrayBuffer
}

async function reparse(blob: Blob) {
  const buf = await blob.arrayBuffer()
  return parseWorkbook(buf)
}

describe('条件格式可编辑(1.9.0):解析保真', () => {
  it('parsed 规则带 id/origin/raw;top10 与 iconSet.reverse 字段解析', async () => {
    const model = await parseWorkbook(await fixtureBuffer())
    const rules = model.sheets[0].conditional
    expect(rules.length).toBe(4)
    for (const r of rules) {
      expect(r.id).toMatch(/^cf-p\d+$/)
      expect(r.origin).toBe('parsed')
      expect(r.raw).toBeTruthy() // 原始 ExcelJS rule 存了 → overlay 原样回写
    }
    const cs = rules.find((r) => r.type === 'colorScale')!
    expect(cs.colorScale).toBeTruthy()
    expect(cs.colorScale!.mid).toBeTruthy() // 3 色标
    const icon = rules.find((r) => r.type === 'iconSet')!
    expect(icon.iconSet?.reverse).toBe(true)
    const t10 = rules.find((r) => r.type === 'top10')!
    expect(t10.top10).toEqual({ rank: 3, percent: false, bottom: true })
  })
})

describe('条件格式可编辑(1.9.0):set-conditional 命令可撤销', () => {
  it('替换 conditional 数组,逆命令还原前态', () => {
    const ruleA: ConditionalRule = { id: 'a', origin: 'user', ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }], priority: 1, type: 'cellIs', operator: 'greaterThan', formulae: ['1'] }
    const ruleB: ConditionalRule = { id: 'b', origin: 'user', ranges: [{ top: 1, left: 0, bottom: 1, right: 0 }], priority: 2, type: 'dataBar', dataBar: { color: '#638EC6', gradient: true } }
    const sheet = { conditional: [ruleA] } as unknown as SheetModel
    const { inverse } = applyCommand(sheet, { kind: 'set-conditional', rules: [ruleA, ruleB] })
    expect(sheet.conditional.map((r) => r.id)).toEqual(['a', 'b'])
    applyCommand(sheet, inverse) // undo
    expect(sheet.conditional.map((r) => r.id)).toEqual(['a'])
  })
})

describe('条件格式可编辑(1.9.0):导出回写往返', () => {
  it('rebuild:原 4 条规则往返存活(parsed raw 原样回写,不退化)', async () => {
    const model = await parseWorkbook(await fixtureBuffer())
    const blob = await workbookToXlsxBlob(model, { fidelity: 'rebuild' })
    const re = await reparse(blob)
    const rules = re.sheets[0].conditional
    expect(rules.some((r) => r.type === 'colorScale')).toBe(true)
    expect(rules.some((r) => r.type === 'cellIs')).toBe(true)
    expect(rules.some((r) => r.type === 'iconSet')).toBe(true)
    expect(rules.some((r) => r.type === 'top10')).toBe(true)
  })

  it('rebuild:用户新建 cellIs 规则导出后往返存活', async () => {
    const model = await parseWorkbook(await fixtureBuffer())
    const sheet = model.sheets[0]
    const userRule: ConditionalRule = {
      id: 'cf-u0', origin: 'user', ranges: [{ top: 0, left: 1, bottom: 5, right: 1 }], priority: 9,
      type: 'cellIs', operator: 'lessThan', formulae: ['20'], style: { font: { bold: true, color: '#FF0000' } },
    }
    sheet.conditional = [...sheet.conditional, userRule]
    const blob = await workbookToXlsxBlob(model, { fidelity: 'rebuild' })
    const re = await reparse(blob)
    const got = re.sheets[0].conditional.filter((r) => r.type === 'cellIs' && r.operator === 'lessThan')
    expect(got.length).toBe(1)
    expect(got[0].formulae).toEqual(['20'])
  })

  it('overlay:原件规则保留 + 用户删一条后导出反映删除', async () => {
    const src = await fixtureBuffer()
    const model = await parseWorkbook(src)
    const sheet = model.sheets[0]
    // 删掉 iconSet 规则(留 3 条)
    sheet.conditional = sheet.conditional.filter((r) => r.type !== 'iconSet')
    const blob = await workbookToXlsxBlob(model, { fidelity: 'overlay', sourceBuffer: src })
    const re = await reparse(blob)
    const rules = re.sheets[0].conditional
    expect(rules.some((r) => r.type === 'iconSet')).toBe(false) // 删除反映了
    expect(rules.some((r) => r.type === 'colorScale')).toBe(true) // 未删的保留
    expect(rules.some((r) => r.type === 'cellIs')).toBe(true)
  })
})
