import { describe, it, expect } from 'vitest'
import { formatValue } from '../number-format'
import { serialToDate } from '../date-serial'

describe('formatValue - 数值', () => {
  it('General 整数/小数', () => {
    expect(formatValue(1234, 'General').text).toBe('1234')
    expect(formatValue(1234.5, 'General').text).toBe('1234.5')
  })

  it('千分位', () => {
    expect(formatValue(1234567, '#,##0').text).toBe('1,234,567')
    expect(formatValue(1234.5, '#,##0.00').text).toBe('1,234.50')
  })

  it('固定小数位', () => {
    expect(formatValue(3.14159, '0.00').text).toBe('3.14')
    expect(formatValue(5, '0.00').text).toBe('5.00')
  })

  it('百分比', () => {
    expect(formatValue(0.1234, '0.00%').text).toBe('12.34%')
    expect(formatValue(0.5, '0%').text).toBe('50%')
  })

  it('负数红括号(四段)', () => {
    const neg = formatValue(-1234, '#,##0;[Red](#,##0)')
    expect(neg.text).toBe('(1,234)')
    expect(neg.color).toBe('#FF0000')
  })

  it('零段', () => {
    expect(formatValue(0, '0.00;-0.00;"零"').text).toBe('零')
  })

  it('单段负数自动补负号', () => {
    expect(formatValue(-42, '0').text).toBe('-42')
  })

  it('条件段 [>=100]', () => {
    const r = formatValue(150, '[>=100]"大";[<100]"小"')
    expect(r.text).toBe('大')
    expect(formatValue(50, '[>=100]"大";[<100]"小"').text).toBe('小')
  })

  it('科学计数', () => {
    expect(formatValue(12345, '0.00E+00').text).toMatch(/1\.23E\+04/)
  })

  it('文本段 @', () => {
    expect(formatValue('hi', '"<"@">"').text).toBe('<hi>')
  })
})

describe('formatValue - 日期', () => {
  it('序列号转日期(1900 系统)', () => {
    // 2020-01-01 的序列号是 43831
    const d = serialToDate(43831)
    expect(d.getUTCFullYear()).toBe(2020)
    expect(d.getUTCMonth()).toBe(0)
    expect(d.getUTCDate()).toBe(1)
  })

  it('1900 闰年 bug: 序列号 60 仍映射到 1900-02-28 之后修正', () => {
    // 序列号 61 = 1900-03-01(因为 60 是不存在的 2-29)
    const d = serialToDate(61)
    expect(d.getUTCMonth()).toBe(2) // March
    expect(d.getUTCDate()).toBe(1)
  })

  it('yyyy-mm-dd 格式化', () => {
    const r = formatValue(43831, 'yyyy-mm-dd')
    expect(r.text).toBe('2020-01-01')
  })

  it('中文日期格式', () => {
    const r = formatValue(43831, 'yyyy"年"m"月"d"日"')
    expect(r.text).toBe('2020年1月1日')
  })

  it('时间格式 h:mm', () => {
    // 0.5 = 12:00
    expect(formatValue(0.5, 'h:mm').text).toBe('12:00')
  })
})
