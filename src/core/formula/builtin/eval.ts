/**
 * 内置公式引擎 —— 求值器(框架无关,纯函数)。1.14.0 新增。
 * 走 AST,引用经 RefResolver 取当前计算值;区域 → Matrix;运算符 + 函数求值,错误就近传播。
 */
import type { Node, BinOp } from './parse'
import { FUNCTIONS } from './functions'
import { type EvalResult, type Scalar, FErr, ERR, isErr, scalarOf, toNum, toStr } from './values'

export interface RefResolver {
  /** 取某格当前计算标量值;sheetName=null=当前表;越界/空 → null。 */
  getCell(sheetName: string | null, row: number, col: number): Scalar
}

export function evalAst(node: Node, res: RefResolver): EvalResult {
  switch (node.t) {
    case 'num': return node.v
    case 'str': return node.v
    case 'bool': return node.v
    case 'err': return new FErr(node.v)
    case 'ref': return res.getCell(node.sheet, node.row, node.col)
    case 'range': {
      const r0 = Math.min(node.a.row, node.b.row), r1 = Math.max(node.a.row, node.b.row)
      const c0 = Math.min(node.a.col, node.b.col), c1 = Math.max(node.a.col, node.b.col)
      const rows: Scalar[][] = []
      for (let r = r0; r <= r1; r++) { const row: Scalar[] = []; for (let c = c0; c <= c1; c++) row.push(res.getCell(node.a.sheet, r, c)); rows.push(row) }
      return { __m: true, rows }
    }
    case 'unary': {
      const v = toNum(scalarOf(evalAst(node.x, res)))
      if (isErr(v)) return v
      if (node.op === '%') return v / 100
      return node.op === '-' ? -v : v
    }
    case 'bin': return evalBin(node.op, evalAst(node.l, res), evalAst(node.r, res))
    case 'func': {
      const fn = FUNCTIONS[node.name]
      if (!fn) return ERR.name()
      const args = node.args.map((a) => evalAst(a, res))
      try { return fn(args) } catch { return ERR.value() }
    }
  }
}

function evalBin(op: BinOp, lv: EvalResult, rv: EvalResult): Scalar {
  const l = scalarOf(lv), r = scalarOf(rv)
  if (isErr(l)) return l
  if (isErr(r)) return r
  if (op === '&') { const a = toStr(l), b = toStr(r); if (isErr(a)) return a; if (isErr(b)) return b; return a + b }
  if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') return compare(op, l, r)
  // 算术
  const a = toNum(l), b = toNum(r)
  if (isErr(a)) return a
  if (isErr(b)) return b
  switch (op) {
    case '+': return a + b
    case '-': return a - b
    case '*': return a * b
    case '/': return b === 0 ? ERR.div0() : a / b
    case '^': { const p = Math.pow(a, b); return isNaN(p) ? ERR.num() : p }
  }
}

function compare(op: BinOp, l: Scalar, r: Scalar): boolean {
  let cmp: number
  if (typeof l === 'number' && typeof r === 'number') cmp = l === r ? 0 : l < r ? -1 : 1
  else if (typeof l === 'boolean' || typeof r === 'boolean') { const a = l === true ? 1 : 0, b = r === true ? 1 : 0; cmp = a - b }
  else { const a = String(l ?? '').toLowerCase(), b = String(r ?? '').toLowerCase(); cmp = a === b ? 0 : a < b ? -1 : 1 }
  switch (op) {
    case '=': return cmp === 0
    case '<>': return cmp !== 0
    case '<': return cmp < 0
    case '>': return cmp > 0
    case '<=': return cmp <= 0
    case '>=': return cmp >= 0
    default: return false
  }
}
