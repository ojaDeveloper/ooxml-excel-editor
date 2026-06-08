/**
 * 数字格式 mini 语言引擎。
 *
 * 覆盖: 四段(正;负;零;文本)、条件段 [>=100]、颜色 [Red]/[Color5]、千分位 ,、
 * 小数 0/#/?、百分比 %、科学计数 E+、缩放(尾随逗号)、字面量/引号/反斜杠转义、
 * 占位 _ 与填充 *(占位渲染为空格)、日期时间 y/m/d/h/s/AM/PM/[h] 经过时间、分数 ?/?。
 *
 * 不追求与 Excel 像素级一致，目标是常见格式肉眼正确。
 */
import { serialToDate } from './date-serial'

export interface FormatResult {
  text: string
  color?: string
}

const COLOR_NAMES: Record<string, string> = {
  black: '#000000',
  blue: '#0000FF',
  cyan: '#00FFFF',
  green: '#008000',
  magenta: '#FF00FF',
  red: '#FF0000',
  white: '#FFFFFF',
  yellow: '#FFFF00',
}
// [Color1..56] 近似映射到 indexed 调色板的一个子集(只挑常见几个)
const COLOR_INDEX: Record<number, string> = {
  1: '#000000', 2: '#FFFFFF', 3: '#FF0000', 4: '#00FF00', 5: '#0000FF',
  6: '#FFFF00', 7: '#FF00FF', 8: '#00FFFF', 10: '#008000', 12: '#000080',
}

interface Section {
  raw: string
  color?: string
  condition?: { op: string; value: number }
  hasDateTime: boolean
}

const sectionCache = new Map<string, Section[]>()

export function formatValue(
  value: number | string | boolean | Date | null,
  code: string | undefined,
  date1904 = false,
): FormatResult {
  if (value === null || value === undefined) return { text: '' }
  if (!code || code === 'General' || code === '@' && typeof value !== 'string') {
    if (code === '@') return { text: String(value) }
    return { text: generalFormat(value) }
  }

  // 字符串值: 用含 @ 的段(优先第 4 段)，否则原样
  if (typeof value === 'string') {
    const sections = parseSections(code)
    const textSec = sections[3] && /@/.test(sections[3].raw)
      ? sections[3]
      : sections.find((s) => /@/.test(s.raw))
    if (textSec) return renderTextSection(textSec, value)
    return { text: value }
  }
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE' }

  // 数值 / 日期
  let num: number
  if (value instanceof Date) {
    num = dateToSerial(value, date1904)
  } else {
    num = value
  }

  const sections = parseSections(code)
  const sec = pickSection(sections, num)
  if (!sec) return { text: generalFormat(num) }

  if (sec.hasDateTime) {
    const dateVal = value instanceof Date ? value : serialToDate(num, date1904)
    return { text: renderDateSection(sec.raw, dateVal, num), color: sec.color }
  }

  // 始终用绝对值渲染量级，符号由调用方决定(避免重复负号)
  const isNeg = num < 0
  const out = renderNumberSection(sec.raw, Math.abs(num))
  if (isNeg && !hasExplicitSign(sec.raw)) {
    return { text: '-' + out, color: sec.color }
  }
  return { text: out, color: sec.color }
}

function hasExplicitSign(raw: string): boolean {
  return /[()-]/.test(stripDecorations(raw))
}

// ---------------- 分段 ----------------
function parseSections(code: string): Section[] {
  const cached = sectionCache.get(code)
  if (cached) return cached
  const parts = splitTopLevel(code, ';')
  const sections = parts.map(parseSection)
  sectionCache.set(code, sections)
  return sections
}

function splitTopLevel(code: string, sep: string): string[] {
  const out: string[] = []
  let cur = ''
  let inQuote = false
  let inBracket = false
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '\\') {
      cur += ch + (code[i + 1] ?? '')
      i++
      continue
    }
    if (ch === '"') {
      inQuote = !inQuote
      cur += ch
      continue
    }
    if (!inQuote && ch === '[') inBracket = true
    if (!inQuote && ch === ']') inBracket = false
    if (ch === sep && !inQuote && !inBracket) {
      out.push(cur)
      cur = ''
      continue
    }
    cur += ch
  }
  out.push(cur)
  return out
}

function parseSection(raw: string): Section {
  let color: string | undefined
  let condition: Section['condition']
  // 提取所有方括号修饰(颜色/条件)，但保留 [h]/[m]/[s] 经过时间标记
  const cleaned = raw.replace(/\[([^\]]+)\]/g, (full, inner: string) => {
    const low = inner.toLowerCase().trim()
    if (/^(h+|m+|s+)$/.test(low)) return full // 经过时间，保留
    if (COLOR_NAMES[low]) {
      color = COLOR_NAMES[low]
      return ''
    }
    const cm = /^color(\d+)$/.exec(low)
    if (cm) {
      color = COLOR_INDEX[+cm[1]] ?? '#000000'
      return ''
    }
    const cond = /^(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d+)?)$/.exec(low)
    if (cond) {
      condition = { op: cond[1], value: parseFloat(cond[2]) }
      return ''
    }
    if (low === 'red' || low === 'blue') {
      color = COLOR_NAMES[low]
      return ''
    }
    return '' // 其它修饰(locale 等)忽略
  })
  return {
    raw: cleaned,
    color,
    condition,
    hasDateTime: detectDateTime(cleaned),
  }
}

function detectDateTime(raw: string): boolean {
  // 去掉引号内字面量后看是否有日期/时间 token
  const stripped = stripQuoted(raw)
  return /(\[?[hms]+\]?|y+|d+|am\/pm|a\/p)/i.test(stripped) &&
    !/E[+-]/.test(stripped) // 排除科学计数里的 e
}

function pickSection(sections: Section[], num: number): Section | undefined {
  // 含条件段: 按顺序匹配条件
  const conditional = sections.filter((s) => s.condition)
  if (conditional.length) {
    for (const s of sections) {
      if (s.condition && matchCondition(s.condition, num)) return s
      if (!s.condition) {
        // 无条件段作为"否则"分支
        return s
      }
    }
  }
  // 默认: 正/负/零/文本
  if (num > 0) return sections[0]
  if (num < 0) return sections[1] ?? sections[0]
  return sections[2] ?? sections[0]
}

function matchCondition(c: { op: string; value: number }, num: number): boolean {
  switch (c.op) {
    case '<': return num < c.value
    case '>': return num > c.value
    case '<=': return num <= c.value
    case '>=': return num >= c.value
    case '=': return num === c.value
    case '<>': return num !== c.value
    default: return false
  }
}

// ---------------- 数值渲染 ----------------
function renderNumberSection(raw: string, num: number): string {
  // 百分比
  let value = num
  const percentCount = countUnquoted(raw, '%')
  if (percentCount > 0) value *= Math.pow(100, percentCount)

  // 缩放: 小数点/数字组后紧跟的逗号，每个 /1000
  const scale = trailingCommaScale(raw)
  if (scale > 0) value /= Math.pow(1000, scale)

  // 科学计数
  if (/[eE][+-]?0+/.test(stripQuoted(raw))) {
    return renderScientific(raw, value)
  }
  // 分数
  if (/[#0]*\s*[?#]+\/[?#0]+/.test(stripQuoted(raw))) {
    return renderFraction(raw, value)
  }
  return renderDecimal(raw, value)
}

function renderDecimal(raw: string, value: number): string {
  // 拆出整数格式 / 小数格式
  const numericMask = extractNumericMask(raw)
  const dotIdx = numericMask.indexOf('.')
  const intMask = dotIdx >= 0 ? numericMask.slice(0, dotIdx) : numericMask
  const fracMask = dotIdx >= 0 ? numericMask.slice(dotIdx + 1) : ''

  const useThousands = intMask.includes(',')
  const fracDigits = (fracMask.match(/[0#?]/g) || []).length

  const rounded = roundTo(value, fracDigits)
  const neg = rounded < 0
  const absVal = Math.abs(rounded)

  let intPart = Math.floor(absVal).toString()
  let fracPart = ''
  if (fracDigits > 0) {
    fracPart = absVal.toFixed(fracDigits).split('.')[1] ?? ''
    // 去掉超出 minFracDigits 且为 # 的尾零
    fracPart = trimFraction(fracPart, fracMask)
  }

  // 整数最少位数(0 的个数)
  const minIntDigits = (intMask.match(/0/g) || []).length
  if (intPart.length < minIntDigits) intPart = intPart.padStart(minIntDigits, '0')
  if (intPart === '0' && (intMask.match(/[0#]/g) || []).length === 0) intPart = ''

  if (useThousands) intPart = addThousands(intPart)

  // 把格式化好的数字回填到原始 token 流(保留字面量/货币符号/占位)
  let body = fillNumberIntoMask(raw, (neg ? '-' : '') + intPart + (fracPart ? '.' + fracPart : ''))
  return body
}

/** 把数字串塞回 mask: 替换第一段连续数字占位为结果，其余字面量保留 */
function fillNumberIntoMask(raw: string, formatted: string): string {
  // 纯字面量(无数字占位)的段: 直接渲染字面量，不注入数字
  const hasPlaceholder = /[0#?]/.test(extractNumericMask(raw))
  let out = ''
  let injected = !hasPlaceholder
  let i = 0
  while (i < raw.length) {
    const ch = raw[i]
    if (ch === '\\') {
      out += raw[i + 1] ?? ''
      i += 2
      continue
    }
    if (ch === '"') {
      const end = raw.indexOf('"', i + 1)
      out += raw.slice(i + 1, end < 0 ? raw.length : end)
      i = end < 0 ? raw.length : end + 1
      continue
    }
    if (ch === '_') {
      out += ' '
      i += 2
      continue
    }
    if (ch === '*') {
      // 填充字符，简化为不填充
      i += 2
      continue
    }
    if (ch === '[') {
      // 经过时间已在 detectDateTime 处理；这里跳过残留
      const end = raw.indexOf(']', i)
      i = end < 0 ? raw.length : end + 1
      continue
    }
    if (ch === '%' || ch === ',') {
      if (ch === '%') out += '%'
      i++
      continue
    }
    if (ch === '0' || ch === '#' || ch === '?' || ch === '.') {
      if (!injected) {
        out += formatted
        injected = true
      }
      // 吞掉这一整段连续的数字占位
      while (i < raw.length && '0#?.,'.includes(raw[i])) i++
      continue
    }
    out += ch
    i++
  }
  if (!injected) out += formatted
  return out
}

function renderScientific(raw: string, value: number): string {
  const m = /([0#.,]+)[eE]([+-]?)(0+)/.exec(stripQuoted(raw))
  const fracDigits = m ? ((m[1].split('.')[1] || '').match(/[0#]/g) || []).length : 2
  const expSign = m ? m[2] : '+'
  let s = value.toExponential(fracDigits)
  // toExponential 给 "1.23e+4"，转成 Excel 风格 "1.23E+04"
  s = s.replace('e', 'E')
  if (expSign === '+') s = s.replace('E', 'E+').replace('E+-', 'E-').replace('E++', 'E+')
  // 指数补零到 mask 长度
  return s.replace(/E([+-])(\d+)/, (_a, sign, d) => `E${sign}${d.padStart(m ? m[3].length : 2, '0')}`)
}

function renderFraction(_raw: string, value: number): string {
  const neg = value < 0
  const abs = Math.abs(value)
  const whole = Math.floor(abs)
  const frac = abs - whole
  if (frac < 1e-9) return (neg ? '-' : '') + whole.toString()
  // 简单连分数逼近(分母 <= 99)
  const [n, d] = approxFraction(frac, 99)
  const sign = neg ? '-' : ''
  if (whole > 0) return `${sign}${whole} ${n}/${d}`
  return `${sign}${n}/${d}`
}

function approxFraction(x: number, maxDen: number): [number, number] {
  let bestN = 0
  let bestD = 1
  let bestErr = Math.abs(x)
  for (let d = 1; d <= maxDen; d++) {
    const n = Math.round(x * d)
    const err = Math.abs(x - n / d)
    if (err < bestErr) {
      bestErr = err
      bestN = n
      bestD = d
      if (err < 1e-9) break
    }
  }
  return [bestN, bestD]
}

// ---------------- 日期时间渲染 ----------------
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

function renderDateSection(raw: string, date: Date, serial: number): string {
  const tokens = tokenizeDate(raw)
  const Y = date.getUTCFullYear()
  const Mo = date.getUTCMonth() + 1
  const D = date.getUTCDate()
  const H = date.getUTCHours()
  const Mi = date.getUTCMinutes()
  const S = date.getUTCSeconds()
  const dow = date.getUTCDay()
  const ampm = tokens.some((t) => /am\/pm|a\/p/i.test(t))
  const hour12 = ((H % 12) || 12)

  let out = ''
  let prevWasHour = false // 仅在遇到字母 token 时更新，分隔符不重置
  for (let i = 0; i < tokens.length; i++) {
    const tk = tokens[i]
    const low = tk.toLowerCase()
    const nextIsSeconds = nextLetterIsSeconds(tokens, i)
    const isLetterToken = /^[a-z[]/.test(low)
    switch (true) {
      case low === 'yyyy': out += String(Y).padStart(4, '0'); break
      case low === 'yy': out += String(Y % 100).padStart(2, '0'); break
      case low === 'mmmm': out += MONTHS[Mo - 1]; break
      case low === 'mmm': out += MONTHS[Mo - 1].slice(0, 3); break
      case low === 'mmmmm': out += MONTHS[Mo - 1][0]; break
      case low === 'mm' && !prevWasHour && !nextIsSeconds: out += String(Mo).padStart(2, '0'); break
      case low === 'm' && !prevWasHour && !nextIsSeconds: out += String(Mo); break
      case low === 'mm': out += String(Mi).padStart(2, '0'); break // 跟在小时/秒前 → 分钟
      case low === 'm': out += String(Mi); break
      case low === 'dddd': out += DAYS[dow]; break
      case low === 'ddd': out += DAYS[dow].slice(0, 3); break
      case low === 'dd': out += String(D).padStart(2, '0'); break
      case low === 'd': out += String(D); break
      case low === 'hh': out += String(ampm ? hour12 : H).padStart(2, '0'); break
      case low === 'h': out += String(ampm ? hour12 : H); break
      case /^\[h+\]$/.test(low): out += String(Math.floor(serial * 24)); break
      case /^\[m+\]$/.test(low): out += String(Math.floor(serial * 24 * 60)); break
      case /^\[s+\]$/.test(low): out += String(Math.floor(serial * 86400)); break
      case low === 'ss': out += String(S).padStart(2, '0'); break
      case low === 's': out += String(S); break
      case /am\/pm/i.test(low): out += H < 12 ? 'AM' : 'PM'; break
      case /a\/p/i.test(low): out += H < 12 ? 'A' : 'P'; break
      default: out += tk
    }
    if (isLetterToken) prevWasHour = /^h+$/.test(low) || /^\[h+\]$/.test(low)
  }
  return out
}

/** 从 i+1 起跳过分隔符，看下一个字母 token 是否为秒(s/ss/[s]) */
function nextLetterIsSeconds(tokens: string[], i: number): boolean {
  for (let k = i + 1; k < tokens.length; k++) {
    const t = tokens[k].toLowerCase()
    if (/^[a-z[]/.test(t)) return /^s+$/.test(t) || /^\[s+\]$/.test(t)
  }
  return false
}

function tokenizeDate(raw: string): string[] {
  const tokens: string[] = []
  let i = 0
  while (i < raw.length) {
    const ch = raw[i]
    if (ch === '\\') {
      tokens.push(raw[i + 1] ?? '')
      i += 2
      continue
    }
    if (ch === '"') {
      const end = raw.indexOf('"', i + 1)
      tokens.push(raw.slice(i + 1, end < 0 ? raw.length : end))
      i = end < 0 ? raw.length : end + 1
      continue
    }
    if (ch === '[') {
      const end = raw.indexOf(']', i)
      tokens.push(raw.slice(i, end < 0 ? raw.length : end + 1))
      i = end < 0 ? raw.length : end + 1
      continue
    }
    if (/[a-z]/i.test(ch)) {
      // AM/PM 特判
      if (raw.slice(i, i + 5).toLowerCase() === 'am/pm') {
        tokens.push(raw.slice(i, i + 5))
        i += 5
        continue
      }
      if (raw.slice(i, i + 3).toLowerCase() === 'a/p') {
        tokens.push(raw.slice(i, i + 3))
        i += 3
        continue
      }
      let j = i
      while (j < raw.length && raw[j].toLowerCase() === ch.toLowerCase()) j++
      tokens.push(raw.slice(i, j))
      i = j
      continue
    }
    tokens.push(ch)
    i++
  }
  return tokens
}

// ---------------- 文本段 ----------------
function renderTextSection(sec: Section, value: string): FormatResult {
  let out = ''
  const raw = sec.raw
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i]
    if (ch === '@') {
      out += value
    } else if (ch === '"') {
      const end = raw.indexOf('"', i + 1)
      out += raw.slice(i + 1, end < 0 ? raw.length : end)
      i = end < 0 ? raw.length : end
    } else if (ch === '\\') {
      out += raw[i + 1] ?? ''
      i++
    } else {
      out += ch
    }
  }
  return { text: out, color: sec.color }
}

// ---------------- 通用工具 ----------------
function generalFormat(value: number | string | boolean | Date): string {
  if (typeof value === 'number') {
    if (Number.isInteger(value)) return String(value)
    // 模拟 Excel General: 最多 ~11 位有效数字
    let s = value.toPrecision(11)
    if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '')
    if (Math.abs(value) >= 1e11 || (Math.abs(value) < 1e-4 && value !== 0)) {
      s = value.toExponential()
    }
    return s
  }
  if (value instanceof Date) return value.toISOString()
  return String(value)
}

function stripQuoted(raw: string): string {
  return raw.replace(/"[^"]*"/g, '').replace(/\\./g, '')
}
function stripDecorations(raw: string): string {
  return stripQuoted(raw).replace(/\[[^\]]*\]/g, '')
}
function extractNumericMask(raw: string): string {
  // 去掉引号字面量、反斜杠转义、方括号、占位 _x *x、%，保留 0 # ? . ,
  let out = ''
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i]
    if (ch === '\\') { i++; continue }
    if (ch === '"') { const e = raw.indexOf('"', i + 1); i = e < 0 ? raw.length : e; continue }
    if (ch === '[') { const e = raw.indexOf(']', i); i = e < 0 ? raw.length : e; continue }
    if (ch === '_' || ch === '*') { i++; continue }
    if ('0#?.,'.includes(ch)) out += ch
  }
  return out
}
function countUnquoted(raw: string, target: string): number {
  return (stripQuoted(raw).match(new RegExp('\\' + target, 'g')) || []).length
}
function trailingCommaScale(raw: string): number {
  const mask = extractNumericMask(raw)
  // 末尾紧跟在数字占位之后的逗号(且其后无数字占位)
  const m = /[0#?](,+)$/.exec(mask)
  return m ? m[1].length : 0
}
function roundTo(value: number, digits: number): number {
  const f = Math.pow(10, digits)
  return Math.round((value + Number.EPSILON) * f) / f
}
function trimFraction(frac: string, fracMask: string): string {
  const maskChars = fracMask.match(/[0#?]/g) || []
  let result = frac.split('')
  for (let i = result.length - 1; i >= 0; i--) {
    if (maskChars[i] === '#' && result[i] === '0') {
      result.pop()
    } else {
      break
    }
  }
  return result.join('')
}
function addThousands(intStr: string): string {
  const neg = intStr.startsWith('-')
  const digits = neg ? intStr.slice(1) : intStr
  const withSep = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  return (neg ? '-' : '') + withSep
}
/** Date → Excel 序列号(导出/往返用). date1904=true 时 epoch=1904-01-01;
 *  1900 系统时跳过 phantom 1900-02-29(序号 ≥ 60 加 1 补回 bug 偏移). */
export function dateToSerial(date: Date, date1904: boolean): number {
  const MS = 86400 * 1000
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31)
  let serial = (date.getTime() - epoch) / MS
  if (!date1904 && serial >= 60) serial += 1
  return serial
}
