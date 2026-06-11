import { describe, it, expect } from 'vitest'
import { unescapeMsoNumFmt, parseMsoNumberFormat } from '../clipboard-html'

// 真实 WPS 复制出的 mso-number-format 值(CSS 转义),验证解回 Excel 格式码
describe('mso-number-format 解析(WPS/Excel 富粘贴数字格式)', () => {
  it('unescapeMsoNumFmt 还原日期/货币/数字格式码', () => {
    expect(unescapeMsoNumFmt('"yyyy/m/d"')).toBe('yyyy/m/d')
    // 货币红负:\0022→" \#→# \;→; \\(→\(
    expect(unescapeMsoNumFmt('"\\0022￥\\0022\\#,\\#\\#0\\.00_)\\;[Red]\\\\(\\0022￥\\0022\\#,\\#\\#0\\.00\\\\)"')).toBe('"￥"#,##0.00_);[Red]\\("￥"#,##0.00\\)')
    expect(unescapeMsoNumFmt('"0\\.00_)\\;[Red]\\\\(0\\.00\\\\)"')).toBe('0.00_);[Red]\\(0.00\\)')
  })

  it('parseMsoNumberFormat 从声明串里取 numFmt;General 跳过;含转义分号不被截断', () => {
    expect(parseMsoNumberFormat('color:#000;mso-number-format:"yyyy/m/d";border:.5pt solid #000')).toBe('yyyy/m/d')
    expect(parseMsoNumberFormat('mso-number-format:General;border:none')).toBeUndefined()
    expect(parseMsoNumberFormat('border:.5pt solid #000')).toBeUndefined()
    // 值里有转义的 \; 不能在那里截断
    expect(parseMsoNumberFormat('mso-number-format:"0\\.00_)\\;[Red]\\\\(0\\.00\\\\)";border:none')).toBe('0.00_);[Red]\\(0.00\\)')
    // CSS 层叠:合并串是 td默认;类;内联 —— 裸 td 默认的 General 在前,类里的真实格式码在后,取最后一条(否则被 General 顶掉丢日期格式)
    expect(parseMsoNumberFormat('vertical-align:middle;mso-number-format:General;mso-number-format:"yyyy/m/d";text-align:center')).toBe('yyyy/m/d')
    // 反过来:只有 td 默认 General、类没覆盖 → 仍是无特殊格式
    expect(parseMsoNumberFormat('vertical-align:middle;mso-number-format:General;text-align:center')).toBeUndefined()
  })
})
