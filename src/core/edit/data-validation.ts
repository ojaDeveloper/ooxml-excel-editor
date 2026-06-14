/**
 * 数据验证引擎(框架无关,纯函数,可单测)。1.8.0 新增。
 *
 * 职责:给定一条 DataValidationRule 和一个待写入的单元格值,判断是否合法 —— 用于"编辑时拦截非法输入"。
 * 值语义跟 model/mutations.ts 的 inferCell 对齐:文本编辑器提交的是字符串,数字串当数值,空串当清空。
 *  - 'stop'(默认):非法 → 硬拒,不写入,弹出错提示;
 *  - 'warning' / 'information':非法 → 软提示(放行写入,只 toast);
 *  - list / whole / decimal / date / time / textLength:按 type + operator 校验;
 *  - custom:不求值,一律放行(避免误拦)。
 */
import type { DataValidationRule, SheetModel } from '../model/types'
import type { CellValue } from '../model/data-access'

export interface ValidationViolation {
  rule: DataValidationRule
  errorStyle: 'stop' | 'warning' | 'information'
  title?: string
  message: string
}

/** 找命中 (row,col) 的数据验证规则(Excel 一格至多一条;取首个包含该格的)。 */
export function findValidationRuleAt(sheet: SheetModel, row: number, col: number): DataValidationRule | undefined {
  const rules = sheet.dataValidationRules
  if (!rules) return undefined
  return rules.find((r) => row >= r.range.top && row <= r.range.bottom && col >= r.range.left && col <= r.range.right)
}

/**
 * 校验一个值是否满足规则。返回 null = 合法(放行);返回 violation = 非法。
 * 空值一律放行(允许清空 —— 拦截清空对预览/编辑组件太粗暴;allowBlank 只表"空值跳过其余校验")。
 */
export function validateCellValue(rule: DataValidationRule, value: CellValue): ValidationViolation | null {
  if (isBlank(value)) return null
  if (rule.type === 'custom') return null // 自定义公式不求值,放行

  let ok = true
  let defaultMsg = ''
  switch (rule.type) {
    case 'list': {
      const s = stringOf(value)
      ok = (rule.options ?? []).includes(s)
      defaultMsg = `只能从下拉列表中选择:${(rule.options ?? []).join('、')}`
      break
    }
    case 'whole': {
      const n = numberOf(value)
      if (n == null || !Number.isInteger(n)) { ok = false; defaultMsg = '请输入整数' }
      else { ok = checkOperator(n, rule); defaultMsg = `请输入整数(${describeConstraint(rule)})` }
      break
    }
    case 'decimal': {
      const n = numberOf(value)
      if (n == null) { ok = false; defaultMsg = '请输入数值' }
      else { ok = checkOperator(n, rule); defaultMsg = `请输入数值(${describeConstraint(rule)})` }
      break
    }
    case 'textLength': {
      const len = stringOf(value).length
      ok = checkOperator(len, rule)
      defaultMsg = `文本长度需${describeConstraint(rule)}`
      break
    }
    case 'date':
    case 'time': {
      const ms = dateMsOf(value)
      if (ms == null) { ok = false; defaultMsg = rule.type === 'date' ? '请输入日期' : '请输入时间' }
      else { ok = checkOperator(ms, rule, true); defaultMsg = `请输入合法的${rule.type === 'date' ? '日期' : '时间'}(${describeConstraint(rule, true)})` }
      break
    }
  }
  if (ok) return null

  const errorStyle = rule.errorStyle ?? 'stop'
  const message = rule.showErrorMessage && rule.error ? rule.error : defaultMsg
  return { rule, errorStyle, title: rule.errorTitle, message }
}

// ====================== 内部 ======================

function isBlank(v: CellValue): boolean {
  return v === null || v === undefined || (typeof v === 'string' && v === '')
}

function stringOf(v: CellValue): string {
  if (v instanceof Date) return v.toISOString()
  return String(v ?? '')
}

/** 跟 inferCell 一致:数字 → 自身;数字串 → Number;其余 → null。 */
function numberOf(v: CellValue): number | null {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null
  if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) return Number(v)
  return null
}

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30) // 序列值 1 = 1900-01-01

function dateMsOf(v: CellValue): number | null {
  if (v instanceof Date) return v.getTime()
  if (typeof v === 'number') return Number.isFinite(v) ? EXCEL_EPOCH_MS + v * 86_400_000 : null
  if (typeof v === 'string' && v.trim() !== '') {
    const t = Date.parse(v)
    return isNaN(t) ? null : t
  }
  return null
}

/** 把操作数解析成可比较的数(asDate 时按日期/序列 → ms)。无法解析 → null。 */
function operand(f: string | number | undefined, asDate: boolean): number | null {
  if (f === undefined) return null
  if (typeof f === 'string' && f.trim() === '') return null // 空操作数 → 无约束(Number('') 会变 0,坑)
  if (asDate) return dateMsOf(f)
  const n = typeof f === 'number' ? f : Number(f)
  return Number.isFinite(n) ? n : null
}

function checkOperator(v: number, rule: DataValidationRule, asDate = false): boolean {
  const a = operand(rule.formulae[0], asDate)
  const b = operand(rule.formulae[1], asDate)
  const op = rule.operator ?? 'between'
  // 约束操作数缺失/解析失败 → 不拦(不在坏约束上误拒用户)
  if (a == null) return true
  switch (op) {
    case 'between': return b == null ? true : v >= a && v <= b
    case 'notBetween': return b == null ? true : v < a || v > b
    case 'equal': return v === a
    case 'notEqual': return v !== a
    case 'greaterThan': return v > a
    case 'lessThan': return v < a
    case 'greaterThanOrEqual': return v >= a
    case 'lessThanOrEqual': return v <= a
    default: return true
  }
}

/** 拼人类可读的约束描述(默认出错提示用)。 */
function describeConstraint(rule: DataValidationRule, asDate = false): string {
  const fmt = (f: string | number | undefined) => {
    if (f === undefined) return ''
    if (asDate) { const ms = dateMsOf(f); return ms == null ? String(f) : new Date(ms).toISOString().slice(0, 10) }
    return String(f)
  }
  const a = fmt(rule.formulae[0])
  const b = fmt(rule.formulae[1])
  switch (rule.operator ?? 'between') {
    case 'between': return `介于 ${a} 和 ${b} 之间`
    case 'notBetween': return `不介于 ${a} 和 ${b} 之间`
    case 'equal': return `等于 ${a}`
    case 'notEqual': return `不等于 ${a}`
    case 'greaterThan': return `大于 ${a}`
    case 'lessThan': return `小于 ${a}`
    case 'greaterThanOrEqual': return `大于等于 ${a}`
    case 'lessThanOrEqual': return `小于等于 ${a}`
    default: return ''
  }
}
