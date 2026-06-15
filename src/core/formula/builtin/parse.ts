/**
 * 内置公式引擎 —— 词法 + 表达式解析(框架无关,纯函数,可单测)。1.14.0 新增。
 *
 * 支持:数字 / 字符串("...") / 布尔(TRUE/FALSE) / 单元格引用($A$1 绝对·相对·混合)/ 区域(A1:B2)/
 *       跨表(Sheet1!A1 / 'My Sheet'!A1)/ 函数调用 / 运算符(优先级见下)/ 括号 / 一元 ±  / 百分号 / 错误字面量。
 * 运算符优先级(低→高):比较(= <> < > <= >=) < 连接(&) < 加减(+ -) < 乘除(* /) < 乘方(^) < 一元(- +) < 百分号(%) < 区域(:)。
 *
 * 输出 AST,引擎缓存到公式格上,改值时只重新求值不重新解析。
 */

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: string }
  | { t: 'ref'; sheet: string | null; row: number; col: number; absRow: boolean; absCol: boolean }
  | { t: 'range'; a: Extract<Node, { t: 'ref' }>; b: Extract<Node, { t: 'ref' }> }
  | { t: 'unary'; op: '-' | '+' | '%'; x: Node }
  | { t: 'bin'; op: BinOp; l: Node; r: Node }
  | { t: 'func'; name: string; args: Node[] }

export type BinOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>='

const ERRORS = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!', '#SPILL!']

// ---------------- 词法 ----------------
type Tok =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'err'; v: string }
  | { k: 'name'; v: string } // 函数名 / TRUE / FALSE(后续判定)
  | { k: 'ref'; v: string } // 单元格引用串(含 $ 和可选 Sheet!)
  | { k: 'op'; v: string }
  | { k: 'lparen' }
  | { k: 'rparen' }
  | { k: 'comma' }

const COL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)/
// Sheet!A1 / 'My Sheet'!A1 前缀
const SHEET_RE = /^(?:'((?:[^']|'')*)'|([A-Za-z_一-龥][A-Za-z0-9_.一-龥]*))!/

function tokenize(input: string): Tok[] {
  const toks: Tok[] = []
  let i = 0
  const s = input
  const n = s.length
  while (i < n) {
    const c = s[i]
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { i++; continue }
    // 字符串
    if (c === '"') {
      let j = i + 1
      let str = ''
      while (j < n) {
        if (s[j] === '"') { if (s[j + 1] === '"') { str += '"'; j += 2; continue } break }
        str += s[j]; j++
      }
      toks.push({ k: 'str', v: str }); i = j + 1; continue
    }
    // 错误字面量
    if (c === '#') {
      const m = ERRORS.find((e) => s.startsWith(e, i))
      if (m) { toks.push({ k: 'err', v: m }); i += m.length; continue }
    }
    // 括号 / 逗号
    if (c === '(') { toks.push({ k: 'lparen' }); i++; continue }
    if (c === ')') { toks.push({ k: 'rparen' }); i++; continue }
    if (c === ',') { toks.push({ k: 'comma' }); i++; continue }
    // 多字符运算符
    if (c === '<' && s[i + 1] === '>') { toks.push({ k: 'op', v: '<>' }); i += 2; continue }
    if (c === '<' && s[i + 1] === '=') { toks.push({ k: 'op', v: '<=' }); i += 2; continue }
    if (c === '>' && s[i + 1] === '=') { toks.push({ k: 'op', v: '>=' }); i += 2; continue }
    if ('+-*/^&=<>%:'.includes(c)) { toks.push({ k: 'op', v: c }); i++; continue }
    // 数字(可带小数 / 科学计数)
    if (/[0-9]/.test(c) || (c === '.' && /[0-9]/.test(s[i + 1] || ''))) {
      const m = /^[0-9]*\.?[0-9]+(?:[eE][+-]?[0-9]+)?/.exec(s.slice(i))!
      toks.push({ k: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue
    }
    // Sheet! 前缀 + 引用 / 引用 / 名称(函数 / 布尔)
    const rest = s.slice(i)
    const sheetM = SHEET_RE.exec(rest)
    let sheetPrefix = ''
    let off = 0
    if (sheetM) { sheetPrefix = rest.slice(0, sheetM[0].length); off = sheetM[0].length }
    const refM = COL_RE.exec(rest.slice(off))
    if (refM && !/[A-Za-z0-9_]/.test(rest[off + refM[0].length] || '')) {
      toks.push({ k: 'ref', v: sheetPrefix + refM[0] }); i += off + refM[0].length; continue
    }
    // 名称(函数名 / TRUE / FALSE);允许字母数字下划线点和中文
    const nameM = /^[A-Za-z_一-龥][A-Za-z0-9_.一-龥]*/.exec(rest)
    if (nameM) { toks.push({ k: 'name', v: nameM[0] }); i += nameM[0].length; continue }
    throw new Error('formula: unexpected char ' + c)
  }
  return toks
}

/** 解析单个引用串(可含 Sheet! 前缀)→ ref 节点。 */
function parseRef(raw: string): Extract<Node, { t: 'ref' }> {
  let sheet: string | null = null
  let rest = raw
  const sm = SHEET_RE.exec(raw)
  if (sm) { sheet = sm[1] != null ? sm[1].replace(/''/g, "'") : sm[2]; rest = raw.slice(sm[0].length) }
  const m = COL_RE.exec(rest)!
  const absCol = m[1] === '$'
  const col = colToIndex(m[2].toUpperCase())
  const absRow = m[3] === '$'
  const row = parseInt(m[4], 10) - 1
  return { t: 'ref', sheet, row, col, absRow, absCol }
}

export function colToIndex(letters: string): number {
  let n = 0
  for (let k = 0; k < letters.length; k++) n = n * 26 + (letters.charCodeAt(k) - 64)
  return n - 1
}

// ---------------- 语法(递归下降)----------------
class Parser {
  private p = 0
  constructor(private toks: Tok[]) {}
  private peek(): Tok | undefined { return this.toks[this.p] }
  private next(): Tok | undefined { return this.toks[this.p++] }
  private eat(k: Tok['k']): Tok { const t = this.next(); if (!t || t.k !== k) throw new Error('formula: expected ' + k); return t }

  parse(): Node {
    const node = this.parseCompare()
    if (this.p < this.toks.length) throw new Error('formula: trailing tokens')
    return node
  }
  private isOp(...vs: string[]): boolean { const t = this.peek(); return !!t && t.k === 'op' && vs.includes(t.v) }
  private parseCompare(): Node {
    let l = this.parseConcat()
    while (this.isOp('=', '<>', '<', '>', '<=', '>=')) { const op = (this.next() as { v: BinOp }).v; l = { t: 'bin', op, l, r: this.parseConcat() } }
    return l
  }
  private parseConcat(): Node {
    let l = this.parseAdd()
    while (this.isOp('&')) { this.next(); l = { t: 'bin', op: '&', l, r: this.parseAdd() } }
    return l
  }
  private parseAdd(): Node {
    let l = this.parseMul()
    while (this.isOp('+', '-')) { const op = (this.next() as { v: '+' | '-' }).v; l = { t: 'bin', op, l, r: this.parseMul() } }
    return l
  }
  private parseMul(): Node {
    let l = this.parsePow()
    while (this.isOp('*', '/')) { const op = (this.next() as { v: '*' | '/' }).v; l = { t: 'bin', op, l, r: this.parsePow() } }
    return l
  }
  private parsePow(): Node {
    const l = this.parseUnary()
    if (this.isOp('^')) { this.next(); return { t: 'bin', op: '^', l, r: this.parsePow() } } // 右结合
    return l
  }
  private parseUnary(): Node {
    if (this.isOp('-', '+')) { const op = (this.next() as { v: '-' | '+' }).v; return { t: 'unary', op, x: this.parseUnary() } }
    return this.parsePostfix()
  }
  private parsePostfix(): Node {
    let x = this.parsePrimary()
    while (this.isOp('%')) { this.next(); x = { t: 'unary', op: '%', x } }
    return x
  }
  private parsePrimary(): Node {
    const t = this.next()
    if (!t) throw new Error('formula: unexpected end')
    if (t.k === 'num') return { t: 'num', v: t.v }
    if (t.k === 'str') return { t: 'str', v: t.v }
    if (t.k === 'err') return { t: 'err', v: t.v }
    if (t.k === 'lparen') { const e = this.parseCompare(); this.eat('rparen'); return e }
    if (t.k === 'ref') {
      const a = parseRef(t.v)
      if (this.isOp(':')) { this.next(); const bt = this.eat('ref') as { v: string }; return { t: 'range', a, b: parseRef(bt.v) } }
      return a
    }
    if (t.k === 'name') {
      const up = t.v.toUpperCase()
      if (this.peek()?.k === 'lparen') { this.next(); const args = this.parseArgs(); this.eat('rparen'); return { t: 'func', name: up, args } }
      if (up === 'TRUE') return { t: 'bool', v: true }
      if (up === 'FALSE') return { t: 'bool', v: false }
      // 裸名当 #NAME?(未定义名称)
      return { t: 'err', v: '#NAME?' }
    }
    throw new Error('formula: unexpected token')
  }
  private parseArgs(): Node[] {
    const args: Node[] = []
    if (this.peek()?.k === 'rparen') return args
    args.push(this.parseCompare())
    while (this.peek()?.k === 'comma') { this.next(); args.push(this.parseCompare()) }
    return args
  }
}

/** 解析公式文本(可带或不带前导 '=')→ AST。解析失败抛错(引擎兜底成 #NAME?/#VALUE?)。 */
export function parseFormula(text: string): Node {
  const body = text[0] === '=' ? text.slice(1) : text
  return new Parser(tokenize(body)).parse()
}
