/**
 * 内置公式引擎 —— 值类型 + 强制转换 + 错误(框架无关,纯函数)。1.14.0 新增。
 * Scalar = 标量;Matrix = 区域求值得到的二维标量数组;FErr = 错误值(#DIV/0! 等)。
 */
export class FErr {
  constructor(public code: string) {}
}
export const ERR = {
  div0: () => new FErr('#DIV/0!'),
  na: () => new FErr('#N/A'),
  name: () => new FErr('#NAME?'),
  num: () => new FErr('#NUM!'),
  ref: () => new FErr('#REF!'),
  value: () => new FErr('#VALUE!'),
}
export function isErr(v: unknown): v is FErr {
  return v instanceof FErr
}

export type Scalar = number | string | boolean | null | FErr
export type Matrix = { __m: true; rows: Scalar[][] }
export type EvalResult = Scalar | Matrix
export function isMatrix(v: EvalResult): v is Matrix {
  return !!v && typeof v === 'object' && (v as Matrix).__m === true
}

/** 区域/数组转标量(标量上下文):取左上;空区 → #VALUE!。 */
export function scalarOf(v: EvalResult): Scalar {
  if (isMatrix(v)) {
    const r = v.rows
    return r.length && r[0].length ? r[0][0] : ERR.value()
  }
  return v
}

/** 把任意 EvalResult 摊平成标量列表(聚合函数用);Matrix → 行优先全部格。 */
export function flatten(v: EvalResult): Scalar[] {
  if (isMatrix(v)) { const out: Scalar[] = []; for (const row of v.rows) for (const c of row) out.push(c); return out }
  return [v]
}

/** 标量 → 数值(算术上下文):number 原样;bool→1/0;null→0;数字串→数;其它→#VALUE!;错误透传。 */
export function toNum(v: Scalar): number | FErr {
  if (isErr(v)) return v
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  if (v == null) return 0
  const t = v.trim()
  if (t === '') return 0
  const n = Number(t)
  return isNaN(n) ? ERR.value() : n
}

/** 标量 → 字符串(连接/文本上下文):number→文本;bool→TRUE/FALSE;null→'';错误透传。 */
export function toStr(v: Scalar): string | FErr {
  if (isErr(v)) return v
  if (v == null) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'number') return numToStr(v)
  return v
}

/** 标量 → 布尔:bool 原样;number→非0;'TRUE'/'FALSE';null→false;其它→#VALUE!。 */
export function toBool(v: Scalar): boolean | FErr {
  if (isErr(v)) return v
  if (typeof v === 'boolean') return v
  if (typeof v === 'number') return v !== 0
  if (v == null) return false
  const u = v.trim().toUpperCase()
  if (u === 'TRUE') return true
  if (u === 'FALSE' || u === '') return false
  return ERR.value()
}

export function numToStr(n: number): string {
  if (!isFinite(n)) return '#NUM!'
  // 跟 General 显示靠近:整数不带小数;否则用 JS 默认(够用)
  return Number.isInteger(n) ? String(n) : String(Number(n.toPrecision(15)))
}
