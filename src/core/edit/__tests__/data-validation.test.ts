import { describe, it, expect } from 'vitest'
import { findValidationRuleAt, validateCellValue } from '../data-validation'
import type { DataValidationRule, SheetModel } from '../../model/types'

const rule = (over: Partial<DataValidationRule>): DataValidationRule => ({
  range: { top: 0, left: 0, bottom: 0, right: 0 },
  type: 'whole',
  formulae: [],
  allowBlank: true,
  ...over,
})

describe('数据验证引擎 validateCellValue', () => {
  it('空值一律放行(允许清空,allowBlank 与否都不拦)', () => {
    expect(validateCellValue(rule({ type: 'whole', operator: 'between', formulae: [1, 10], allowBlank: false }), '')).toBeNull()
    expect(validateCellValue(rule({ type: 'list', options: ['a'], allowBlank: false }), null)).toBeNull()
  })

  it('list:必须命中选项,否则拦', () => {
    const r = rule({ type: 'list', options: ['苹果', '香蕉'] })
    expect(validateCellValue(r, '苹果')).toBeNull()
    const v = validateCellValue(r, '梨')
    expect(v).toBeTruthy()
    expect(v!.errorStyle).toBe('stop')
  })

  it('whole between:整数且在范围内', () => {
    const r = rule({ type: 'whole', operator: 'between', formulae: [10, 100] })
    expect(validateCellValue(r, '50')).toBeNull()
    expect(validateCellValue(r, 50)).toBeNull()
    expect(validateCellValue(r, '5')).toBeTruthy() // 越下界
    expect(validateCellValue(r, '200')).toBeTruthy() // 越上界
    expect(validateCellValue(r, '50.5')).toBeTruthy() // 非整数
    expect(validateCellValue(r, 'abc')).toBeTruthy() // 非数值
  })

  it('decimal greaterThan:数值且大于', () => {
    const r = rule({ type: 'decimal', operator: 'greaterThan', formulae: [0] })
    expect(validateCellValue(r, '0.01')).toBeNull()
    expect(validateCellValue(r, '0')).toBeTruthy()
    expect(validateCellValue(r, '-1')).toBeTruthy()
  })

  it('textLength lessThanOrEqual:长度受限', () => {
    const r = rule({ type: 'textLength', operator: 'lessThanOrEqual', formulae: [5] })
    expect(validateCellValue(r, 'hello')).toBeNull()
    expect(validateCellValue(r, 'hello!')).toBeTruthy()
  })

  it('date between:ISO 串约束,输入日期串比较', () => {
    const r = rule({ type: 'date', operator: 'between', formulae: ['2026-01-01T00:00:00.000Z', '2026-12-31T00:00:00.000Z'] })
    expect(validateCellValue(r, '2026-06-15')).toBeNull()
    expect(validateCellValue(r, '2025-12-31')).toBeTruthy()
    expect(validateCellValue(r, '2027-01-01')).toBeTruthy()
  })

  it('custom:不求值,放行', () => {
    expect(validateCellValue(rule({ type: 'custom', formulae: ['=A1>0'] }), 'anything')).toBeNull()
  })

  it('errorStyle warning/information:仍是 violation,但调用方据此放行', () => {
    const w = validateCellValue(rule({ type: 'whole', operator: 'greaterThan', formulae: [0], errorStyle: 'warning' }), '-1')
    expect(w!.errorStyle).toBe('warning')
  })

  it('自定义出错信息:showErrorMessage + error 优先于默认文案', () => {
    const r = rule({ type: 'whole', operator: 'greaterThan', formulae: [0], showErrorMessage: true, error: '数量必须为正', errorTitle: '非法' })
    const v = validateCellValue(r, '-5')
    expect(v!.message).toBe('数量必须为正')
    expect(v!.title).toBe('非法')
  })

  it('坏约束(操作数解析失败)不误拦', () => {
    const r = rule({ type: 'whole', operator: 'between', formulae: ['', ''] })
    expect(validateCellValue(r, '7')).toBeNull()
  })
})

describe('findValidationRuleAt', () => {
  const sheet = {
    dataValidationRules: [
      { range: { top: 1, left: 1, bottom: 3, right: 3 }, type: 'whole', formulae: [0], allowBlank: true },
    ],
  } as unknown as SheetModel

  it('命中区域内返回规则,区域外返回 undefined', () => {
    expect(findValidationRuleAt(sheet, 2, 2)?.type).toBe('whole')
    expect(findValidationRuleAt(sheet, 1, 1)?.type).toBe('whole') // 边界
    expect(findValidationRuleAt(sheet, 0, 0)).toBeUndefined()
    expect(findValidationRuleAt(sheet, 4, 4)).toBeUndefined()
  })

  it('无规则表返回 undefined', () => {
    expect(findValidationRuleAt({} as SheetModel, 0, 0)).toBeUndefined()
  })
})
