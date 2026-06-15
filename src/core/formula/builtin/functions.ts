/**
 * 内置公式引擎 —— 函数库(框架无关,纯函数)。1.14.0 新增。覆盖日常常用 ~60 个;按需易扩。
 * 约定:每个函数收已求值的参数(EvalResult[]),返回 Scalar。错误就近传播。
 */
import { type EvalResult, type Scalar, FErr, ERR, isErr, isMatrix, flatten, scalarOf, toNum, toStr, toBool, numToStr } from './values'

export type FnImpl = (args: EvalResult[]) => Scalar

// ---------- 取数辅助 ----------
function nums(args: EvalResult[]): number[] | FErr {
  const out: number[] = []
  for (const a of args) for (const s of flatten(a)) {
    if (isErr(s)) return s
    if (typeof s === 'number') out.push(s)
  }
  return out
}
function num1(v: EvalResult): number | FErr { return toNum(scalarOf(v)) }
function str1(v: EvalResult): string | FErr { return toStr(scalarOf(v)) }
function int1(v: EvalResult): number | FErr { const n = num1(v); return isErr(n) ? n : Math.trunc(n) }

// ---------- 数值/聚合 ----------
const SUM: FnImpl = (a) => { const n = nums(a); return isErr(n) ? n : n.reduce((x, y) => x + y, 0) }
const PRODUCT: FnImpl = (a) => { const n = nums(a); return isErr(n) ? n : n.reduce((x, y) => x * y, 1) }
const AVERAGE: FnImpl = (a) => { const n = nums(a); if (isErr(n)) return n; return n.length ? n.reduce((x, y) => x + y, 0) / n.length : ERR.div0() }
const MAX: FnImpl = (a) => { const n = nums(a); if (isErr(n)) return n; return n.length ? Math.max(...n) : 0 }
const MIN: FnImpl = (a) => { const n = nums(a); if (isErr(n)) return n; return n.length ? Math.min(...n) : 0 }
const COUNT: FnImpl = (a) => nums(a) instanceof FErr ? (nums(a) as FErr) : (nums(a) as number[]).length
const COUNTA: FnImpl = (a) => { let c = 0; for (const x of a) for (const s of flatten(x)) if (s !== null && s !== '') c++; return c }
const COUNTBLANK: FnImpl = (a) => { let c = 0; for (const x of a) for (const s of flatten(x)) if (s === null || s === '') c++; return c }

const round = (fn: (n: number) => number): FnImpl => (a) => {
  const n = num1(a[0]); if (isErr(n)) return n
  const d = a[1] != null ? num1(a[1]) : 0; if (isErr(d)) return d
  const f = Math.pow(10, d); return fn(n * f) / f
}
const ABS: FnImpl = (a) => { const n = num1(a[0]); return isErr(n) ? n : Math.abs(n) }
const INT: FnImpl = (a) => { const n = num1(a[0]); return isErr(n) ? n : Math.floor(n) }
const MOD: FnImpl = (a) => { const x = num1(a[0]), y = num1(a[1]); if (isErr(x)) return x; if (isErr(y)) return y; return y === 0 ? ERR.div0() : x - Math.floor(x / y) * y }
const SQRT: FnImpl = (a) => { const n = num1(a[0]); if (isErr(n)) return n; return n < 0 ? ERR.num() : Math.sqrt(n) }
const POWER: FnImpl = (a) => { const x = num1(a[0]), y = num1(a[1]); if (isErr(x)) return x; if (isErr(y)) return y; return Math.pow(x, y) }
const SIGN: FnImpl = (a) => { const n = num1(a[0]); return isErr(n) ? n : Math.sign(n) }
const CEILING: FnImpl = (a) => { const n = num1(a[0]); if (isErr(n)) return n; const s = a[1] != null ? num1(a[1]) : 1; if (isErr(s)) return s; return s === 0 ? 0 : Math.ceil(n / s) * s }
const FLOOR: FnImpl = (a) => { const n = num1(a[0]); if (isErr(n)) return n; const s = a[1] != null ? num1(a[1]) : 1; if (isErr(s)) return s; return s === 0 ? ERR.div0() : Math.floor(n / s) * s }
const SUMPRODUCT: FnImpl = (a) => {
  const cols = a.map((x) => flatten(x).map((s) => (typeof s === 'number' ? s : isErr(s) ? s : 0)))
  const len = cols[0]?.length ?? 0
  let total = 0
  for (let i = 0; i < len; i++) { let p = 1; for (const c of cols) { const v = c[i]; if (isErr(v)) return v; p *= v as number } total += p }
  return total
}

// ---------- 逻辑 ----------
const IF: FnImpl = (a) => { const c = toBool(scalarOf(a[0])); if (isErr(c)) return c; return c ? scalarOf(a[1] ?? false) : scalarOf(a[2] ?? false) }
const IFERROR: FnImpl = (a) => { const v = scalarOf(a[0]); return isErr(v) ? scalarOf(a[1] ?? '') : v }
const IFNA: FnImpl = (a) => { const v = scalarOf(a[0]); return isErr(v) && v.code === '#N/A' ? scalarOf(a[1] ?? '') : v }
const AND: FnImpl = (a) => { let r = true; for (const x of a) for (const s of flatten(x)) { if (isErr(s)) return s; if (s == null) continue; const b = toBool(s); if (isErr(b)) return b; r = r && b } return r }
const OR: FnImpl = (a) => { let r = false; for (const x of a) for (const s of flatten(x)) { if (isErr(s)) return s; if (s == null) continue; const b = toBool(s); if (isErr(b)) return b; r = r || b } return r }
const NOT: FnImpl = (a) => { const b = toBool(scalarOf(a[0])); return isErr(b) ? b : !b }
const XOR: FnImpl = (a) => { let c = 0; for (const x of a) for (const s of flatten(x)) { if (isErr(s)) return s; if (s == null) continue; const b = toBool(s); if (isErr(b)) return b; if (b) c++ } return c % 2 === 1 }
const IFS: FnImpl = (a) => { for (let i = 0; i + 1 < a.length; i += 2) { const c = toBool(scalarOf(a[i])); if (isErr(c)) return c; if (c) return scalarOf(a[i + 1]) } return ERR.na() }

// ---------- 文本 ----------
const CONCAT: FnImpl = (a) => { let s = ''; for (const x of a) for (const v of flatten(x)) { const t = toStr(v); if (isErr(t)) return t; s += t } return s }
const LEN: FnImpl = (a) => { const s = str1(a[0]); return isErr(s) ? s : s.length }
const LEFT: FnImpl = (a) => { const s = str1(a[0]); if (isErr(s)) return s; const n = a[1] != null ? num1(a[1]) : 1; if (isErr(n)) return n; return s.slice(0, Math.max(0, n)) }
const RIGHT: FnImpl = (a) => { const s = str1(a[0]); if (isErr(s)) return s; const n = a[1] != null ? num1(a[1]) : 1; if (isErr(n)) return n; return n <= 0 ? '' : s.slice(-n) }
const MID: FnImpl = (a) => { const s = str1(a[0]); if (isErr(s)) return s; const st = num1(a[1]), ln = num1(a[2]); if (isErr(st)) return st; if (isErr(ln)) return ln; return s.substr(Math.max(0, st - 1), Math.max(0, ln)) }
const LOWER: FnImpl = (a) => { const s = str1(a[0]); return isErr(s) ? s : s.toLowerCase() }
const UPPER: FnImpl = (a) => { const s = str1(a[0]); return isErr(s) ? s : s.toUpperCase() }
const TRIM: FnImpl = (a) => { const s = str1(a[0]); return isErr(s) ? s : s.replace(/\s+/g, ' ').trim() }
const PROPER: FnImpl = (a) => { const s = str1(a[0]); return isErr(s) ? s : s.replace(/\b\w/g, (c) => c.toUpperCase()).replace(/\B\w/g, (c) => c.toLowerCase()) }
const REPT: FnImpl = (a) => { const s = str1(a[0]); if (isErr(s)) return s; const n = num1(a[1]); if (isErr(n)) return n; return n <= 0 ? '' : s.repeat(Math.floor(n)) }
const EXACT: FnImpl = (a) => { const x = str1(a[0]), y = str1(a[1]); if (isErr(x)) return x; if (isErr(y)) return y; return x === y }
const SUBSTITUTE: FnImpl = (a) => { const s = str1(a[0]), o = str1(a[1]), nw = str1(a[2]); if (isErr(s)) return s; if (isErr(o)) return o; if (isErr(nw)) return nw; if (o === '') return s; return s.split(o).join(nw) }
const FIND: FnImpl = (a) => { const sub = str1(a[0]), s = str1(a[1]); if (isErr(sub)) return sub; if (isErr(s)) return s; const start = a[2] != null ? num1(a[2]) : 1; if (isErr(start)) return start; const i = s.indexOf(sub, start - 1); return i < 0 ? ERR.value() : i + 1 }
const SEARCH: FnImpl = (a) => { const sub = str1(a[0]), s = str1(a[1]); if (isErr(sub)) return sub; if (isErr(s)) return s; const start = a[2] != null ? num1(a[2]) : 1; if (isErr(start)) return start; const i = s.toLowerCase().indexOf(sub.toLowerCase(), start - 1); return i < 0 ? ERR.value() : i + 1 }
const VALUE: FnImpl = (a) => { const n = num1(a[0]); return n }
const TEXTFN: FnImpl = (a) => { const s = str1(a[0]); return isErr(s) ? s : s } // TEXT(简化:不做格式串,返回原文本)

// ---------- 信息 ----------
const ISNUMBER: FnImpl = (a) => typeof scalarOf(a[0]) === 'number'
const ISTEXT: FnImpl = (a) => typeof scalarOf(a[0]) === 'string'
const ISBLANK: FnImpl = (a) => scalarOf(a[0]) === null
const ISLOGICAL: FnImpl = (a) => typeof scalarOf(a[0]) === 'boolean'
const ISERROR: FnImpl = (a) => isErr(scalarOf(a[0]))
const ISNA: FnImpl = (a) => { const v = scalarOf(a[0]); return isErr(v) && v.code === '#N/A' }
const ISEVEN: FnImpl = (a) => { const n = int1(a[0]); return isErr(n) ? n : n % 2 === 0 }
const ISODD: FnImpl = (a) => { const n = int1(a[0]); return isErr(n) ? n : Math.abs(n % 2) === 1 }
const NA: FnImpl = () => ERR.na()

// ---------- 查找 ----------
function asMatrix(v: EvalResult): Scalar[][] { return isMatrix(v) ? v.rows : [[v as Scalar]] }
const matchVal = (target: Scalar, v: Scalar): boolean => {
  if (typeof target === 'number' && typeof v === 'number') return target === v
  return String(target ?? '').toLowerCase() === String(v ?? '').toLowerCase()
}
const VLOOKUP: FnImpl = (a) => {
  const key = scalarOf(a[0]); if (isErr(key)) return key
  const tbl = asMatrix(a[1])
  const colIdx = num1(a[2]); if (isErr(colIdx)) return colIdx
  // 第 4 参 range_lookup:TRUE/省略 = 近似(升序,取最大的 <=key);FALSE = 精确
  const rangeLookup = a[3] != null ? toBool(scalarOf(a[3])) : true
  if (isErr(rangeLookup)) return rangeLookup
  if (!rangeLookup) {
    for (const row of tbl) if (matchVal(key, row[0])) { const v = row[colIdx - 1]; return v === undefined ? ERR.ref() : v }
    return ERR.na()
  }
  let best: Scalar[] | null = null
  for (const row of tbl) { const cell = row[0]; if (typeof key === 'number' && typeof cell === 'number' && cell <= key) best = row; else if (typeof key !== 'number' && String(cell ?? '') <= String(key ?? '')) best = row }
  if (!best) return ERR.na()
  const v = best[colIdx - 1]
  return v === undefined ? ERR.ref() : v
}
const HLOOKUP: FnImpl = (a) => {
  const key = scalarOf(a[0]); if (isErr(key)) return key
  const tbl = asMatrix(a[1])
  const rowIdx = num1(a[2]); if (isErr(rowIdx)) return rowIdx
  const header = tbl[0] ?? []
  for (let c = 0; c < header.length; c++) if (matchVal(key, header[c])) { const v = tbl[rowIdx - 1]?.[c]; return v === undefined ? ERR.ref() : v }
  return ERR.na()
}
const INDEX: FnImpl = (a) => {
  const tbl = asMatrix(a[0])
  const rr = a[1] != null ? num1(a[1]) : 0; if (isErr(rr)) return rr
  const cc = a[2] != null ? num1(a[2]) : 0; if (isErr(cc)) return cc
  if (rr === 0 && tbl.length === 1) { const v = tbl[0][cc - 1]; return v === undefined ? ERR.ref() : v }
  if (cc === 0 && (tbl[0]?.length ?? 0) === 1) { const v = tbl[rr - 1]?.[0]; return v === undefined ? ERR.ref() : v }
  const v = tbl[rr - 1]?.[cc - 1]
  return v === undefined ? ERR.ref() : v
}
const MATCH: FnImpl = (a) => {
  const key = scalarOf(a[0]); if (isErr(key)) return key
  const arr = flatten(a[1])
  const type = a[2] != null ? num1(a[2]) : 1; if (isErr(type)) return type
  if (type === 0) { for (let i = 0; i < arr.length; i++) if (matchVal(key, arr[i])) return i + 1; return ERR.na() }
  // 1 = 最大的 <= key(升序);-1 = 最小的 >= key(降序)
  let best = -1
  for (let i = 0; i < arr.length; i++) { const v = arr[i]; if (typeof v !== 'number' || typeof key !== 'number') continue; if (type === 1 ? v <= key : v >= key) best = i }
  return best < 0 ? ERR.na() : best + 1
}
const CHOOSE: FnImpl = (a) => { const i = num1(a[0]); if (isErr(i)) return i; const v = a[i]; return v === undefined ? ERR.value() : scalarOf(v) }
const ROWS: FnImpl = (a) => asMatrix(a[0]).length
const COLUMNS: FnImpl = (a) => asMatrix(a[0])[0]?.length ?? 0

// ---------- 条件聚合 ----------
function criteria(crit: Scalar): (v: Scalar) => boolean {
  if (typeof crit === 'string') {
    const m = /^(<=|>=|<>|<|>|=)(.*)$/.exec(crit)
    if (m) {
      const op = m[1]; const rhsRaw = m[2]; const rhsNum = Number(rhsRaw)
      const numeric = rhsRaw.trim() !== '' && !isNaN(rhsNum)
      return (v) => {
        if (numeric && typeof v === 'number') {
          switch (op) { case '<': return v < rhsNum; case '>': return v > rhsNum; case '<=': return v <= rhsNum; case '>=': return v >= rhsNum; case '<>': return v !== rhsNum; default: return v === rhsNum }
        }
        const sv = String(v ?? '').toLowerCase(); const sr = rhsRaw.toLowerCase()
        return op === '<>' ? sv !== sr : op === '=' ? sv === sr : false
      }
    }
  }
  // 无运算符:相等匹配(数值或文本,文本不区分大小写)
  return (v) => matchVal(crit, v)
}
function sumifCore(a: EvalResult[], mode: 'sum' | 'count' | 'avg'): Scalar {
  const rng = flatten(a[0])
  const test = criteria(scalarOf(a[1]))
  const sumRng = a[2] != null ? flatten(a[2]) : rng
  let sum = 0, cnt = 0
  for (let i = 0; i < rng.length; i++) {
    if (isErr(rng[i])) return rng[i]
    if (!test(rng[i])) continue
    cnt++
    const sv = sumRng[i]
    if (typeof sv === 'number') sum += sv
  }
  if (mode === 'count') return cnt
  if (mode === 'avg') return cnt ? sum / cnt : ERR.div0()
  return sum
}
const SUMIF: FnImpl = (a) => sumifCore(a, 'sum')
const COUNTIF: FnImpl = (a) => sumifCore([a[0], a[1], a[0]], 'count')
const AVERAGEIF: FnImpl = (a) => sumifCore(a, 'avg')

// ---------- 日期(Excel 序列值;1900 历法,忽略 1900 闰年 bug 对 >=1900-03 的日期无影响)----------
const EPOCH = Date.UTC(1899, 11, 30)
const DAY_MS = 86_400_000
function serialOf(y: number, m: number, d: number): number { return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS) }
function dateOfSerial(serial: number): Date { return new Date(EPOCH + Math.round(serial) * DAY_MS) }
function todaySerial(): number { const n = new Date(); return serialOf(n.getFullYear(), n.getMonth() + 1, n.getDate()) }
const TODAY: FnImpl = () => todaySerial()
const NOW: FnImpl = () => { const n = new Date(); return todaySerial() + (n.getHours() * 3600 + n.getMinutes() * 60 + n.getSeconds()) / 86400 }
const DATE: FnImpl = (a) => { const y = num1(a[0]), m = num1(a[1]), d = num1(a[2]); if (isErr(y)) return y; if (isErr(m)) return m; if (isErr(d)) return d; return serialOf(y, m, d) }
const YEAR: FnImpl = (a) => { const n = num1(a[0]); return isErr(n) ? n : dateOfSerial(n).getUTCFullYear() }
const MONTH: FnImpl = (a) => { const n = num1(a[0]); return isErr(n) ? n : dateOfSerial(n).getUTCMonth() + 1 }
const DAY: FnImpl = (a) => { const n = num1(a[0]); return isErr(n) ? n : dateOfSerial(n).getUTCDate() }
const WEEKDAY: FnImpl = (a) => { const n = num1(a[0]); if (isErr(n)) return n; return dateOfSerial(n).getUTCDay() + 1 }
const DAYS: FnImpl = (a) => { const e = num1(a[0]), s = num1(a[1]); if (isErr(e)) return e; if (isErr(s)) return s; return Math.round(e - s) }

export const FUNCTIONS: Record<string, FnImpl> = {
  SUM, PRODUCT, AVERAGE, MAX, MIN, COUNT, COUNTA, COUNTBLANK,
  ROUND: round(Math.round), ROUNDUP: round((n) => (n >= 0 ? Math.ceil(n) : Math.floor(n))), ROUNDDOWN: round(Math.trunc), TRUNC: round(Math.trunc),
  ABS, INT, MOD, SQRT, POWER, SIGN, CEILING, FLOOR, SUMPRODUCT,
  IF, IFERROR, IFNA, AND, OR, NOT, XOR, IFS,
  CONCAT, CONCATENATE: CONCAT, LEN, LEFT, RIGHT, MID, LOWER, UPPER, TRIM, PROPER, REPT, EXACT, SUBSTITUTE, FIND, SEARCH, VALUE, TEXT: TEXTFN,
  ISNUMBER, ISTEXT, ISBLANK, ISLOGICAL, ISERROR, ISNA, ISEVEN, ISODD, NA, TRUE: () => true, FALSE: () => false,
  VLOOKUP, HLOOKUP, INDEX, MATCH, CHOOSE, ROWS, COLUMNS,
  SUMIF, COUNTIF, AVERAGEIF,
  TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, DAYS,
}

/** 自动补全用:已支持的函数名(升序)。 */
export const FUNCTION_NAMES: string[] = Object.keys(FUNCTIONS).sort()

/** 自动补全参数提示:常用函数的签名;未列出的回退 `NAME(…)`。 */
export const FUNCTION_SIGNATURES: Record<string, string> = {
  SUM: 'SUM(数值1, 数值2, …)', AVERAGE: 'AVERAGE(数值1, …)', COUNT: 'COUNT(值1, …)', COUNTA: 'COUNTA(值1, …)',
  MAX: 'MAX(数值1, …)', MIN: 'MIN(数值1, …)', PRODUCT: 'PRODUCT(数值1, …)', SUMPRODUCT: 'SUMPRODUCT(数组1, 数组2, …)',
  IF: 'IF(条件, 真值, 假值)', IFERROR: 'IFERROR(值, 出错时值)', IFNA: 'IFNA(值, NA时值)', IFS: 'IFS(条件1, 值1, …)',
  AND: 'AND(逻辑1, …)', OR: 'OR(逻辑1, …)', NOT: 'NOT(逻辑)',
  ROUND: 'ROUND(数值, 位数)', ROUNDUP: 'ROUNDUP(数值, 位数)', ROUNDDOWN: 'ROUNDDOWN(数值, 位数)',
  ABS: 'ABS(数值)', INT: 'INT(数值)', MOD: 'MOD(数值, 除数)', SQRT: 'SQRT(数值)', POWER: 'POWER(底, 指数)', CEILING: 'CEILING(数值, [基数])', FLOOR: 'FLOOR(数值, [基数])',
  VLOOKUP: 'VLOOKUP(查找值, 表区域, 列号, [精确])', HLOOKUP: 'HLOOKUP(查找值, 表区域, 行号)', INDEX: 'INDEX(区域, 行号, [列号])', MATCH: 'MATCH(查找值, 区域, [类型])', CHOOSE: 'CHOOSE(序号, 值1, …)',
  SUMIF: 'SUMIF(区域, 条件, [求和区域])', COUNTIF: 'COUNTIF(区域, 条件)', AVERAGEIF: 'AVERAGEIF(区域, 条件, [平均区域])',
  LEFT: 'LEFT(文本, [字数])', RIGHT: 'RIGHT(文本, [字数])', MID: 'MID(文本, 起始, 字数)', LEN: 'LEN(文本)', CONCAT: 'CONCAT(文本1, …)', CONCATENATE: 'CONCATENATE(文本1, …)',
  UPPER: 'UPPER(文本)', LOWER: 'LOWER(文本)', TRIM: 'TRIM(文本)', SUBSTITUTE: 'SUBSTITUTE(文本, 旧, 新)', FIND: 'FIND(子串, 文本, [起始])', SEARCH: 'SEARCH(子串, 文本, [起始])',
  DATE: 'DATE(年, 月, 日)', YEAR: 'YEAR(序列值)', MONTH: 'MONTH(序列值)', DAY: 'DAY(序列值)', TODAY: 'TODAY()', NOW: 'NOW()', WEEKDAY: 'WEEKDAY(序列值)', DAYS: 'DAYS(结束, 开始)',
  ISNUMBER: 'ISNUMBER(值)', ISTEXT: 'ISTEXT(值)', ISBLANK: 'ISBLANK(值)', ISERROR: 'ISERROR(值)', ISNA: 'ISNA(值)',
}
void numToStr
