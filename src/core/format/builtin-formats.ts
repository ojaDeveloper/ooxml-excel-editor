/**
 * 内置数字格式 numFmtId → 格式代码字符串(ECMA-376 §18.8.30)。
 * id 0-49 是保留内置格式；>=164 是自定义格式(从 styles.xml 的 numFmts 取)。
 * 注: 14-22、45-47 等会随 locale 变，这里用通用形态。
 */
export const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '$#,##0;\\-$#,##0',
  6: '$#,##0;[Red]\\-$#,##0',
  7: '$#,##0.00;\\-$#,##0.00',
  8: '$#,##0.00;[Red]\\-$#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0;(#,##0)',
  38: '#,##0;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
}

export function builtinFormat(id: number): string | undefined {
  return BUILTIN_FORMATS[id]
}
