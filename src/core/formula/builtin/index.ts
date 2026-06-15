/**
 * 内置公式引擎(MIT,零依赖)—— 实现 FormulaEngine 接口:解析 + 求值 + 依赖图 + 拓扑级联重算 + 循环检测。1.14.0 新增。
 * 默认引擎(替代 GPL 的 HyperFormula);覆盖日常 ~60 函数(见 functions.ts)。需更全覆盖可注入 HyperFormula。
 *
 * 级联:setCell 改一格 → 闭包出所有(传递)依赖它的公式格 → 递归求值(memo 保证拓扑序)→ 返回值变了的格。
 * 循环引用 → 该环上的格得 #REF!(对齐 Excel"循环引用"近似)。
 */
import type { WorkbookModel } from '../../model/types'
import type { CellValue } from '../../model/data-access'
import type { DirtyCell, FormulaEngine } from '../engine'
import { parseFormula, type Node } from './parse'
import { evalAst, type RefResolver } from './eval'
import { type Scalar, FErr, ERR, isErr } from './values'
import type { FormulaEngineFactory } from '../engine'

/** 内置引擎工厂(MIT,零依赖,同步)。默认引擎。 */
export const builtinFormulaEngineFactory: FormulaEngineFactory = async () => new BuiltinFormulaEngine()

interface EngCell {
  formula?: string // 原始公式串(含 '=')
  ast?: Node | null // 解析后的 AST;null = 解析失败
  value: Scalar // 当前计算值(公式)或字面值
  prec?: Prec[] // 该公式读取的区域(依赖图用)
}
interface Prec { sheet: number; r0: number; c0: number; r1: number; c1: number }

const KEY = (s: number, r: number, c: number) => `${s}:${r}:${c}`
const RANGE_CAP = 100_000 // 单个区域展开进依赖图的格数上限(防 A:A 整列爆内存)

export class BuiltinFormulaEngine implements FormulaEngine {
  private sheets: { name: string; cells: Map<string, EngCell>; rows: number; cols: number }[] = []
  private nameToIndex = new Map<string, number>()
  private dependents = new Map<string, Set<string>>() // 前驱格 key → 依赖它的公式格 key 集合

  setSheets(wb: WorkbookModel): void {
    this.sheets = []
    this.nameToIndex.clear()
    this.dependents.clear()
    // 先登记全部表名(跨表引用到后面的表也能解析)
    for (let i = 0; i < wb.sheets.length; i++) this.nameToIndex.set(wb.sheets[i].name || `Sheet${i + 1}`, i)
    for (let i = 0; i < wb.sheets.length; i++) {
      const s = wb.sheets[i]
      const cells = new Map<string, EngCell>()
      for (const cell of s.cells.values()) {
        const content = cellModelContent(cell)
        if (content == null) continue
        cells.set(KEY(i, cell.row, cell.col), this.makeCell(i, cell.row, cell.col, content, false))
      }
      this.sheets.push({ name: wb.sheets[i].name || `Sheet${i + 1}`, cells, rows: Math.max(s.dimension.rows, 1), cols: Math.max(s.dimension.cols, 1) })
    }
    // 注册依赖 + 初始全量求值(所有公式格)
    const allFormulas: string[] = []
    for (let i = 0; i < this.sheets.length; i++) for (const [k, c] of this.sheets[i].cells) if (c.ast !== undefined) { this.addDeps(k, c.prec); allFormulas.push(k) }
    this.recompute(allFormulas)
  }

  setCell(sheet: number, row: number, col: number, content: string | number | boolean | null): DirtyCell[] {
    const sh = this.sheets[sheet]
    if (!sh) return []
    const key = KEY(sheet, row, col)
    const old = sh.cells.get(key)
    if (old?.prec) this.removeDeps(key, old.prec) // 撤旧依赖
    if (content == null || content === '') {
      sh.cells.delete(key)
    } else {
      const c = this.makeCell(sheet, row, col, content, true)
      sh.cells.set(key, c)
      if (c.ast !== undefined) this.addDeps(key, c.prec)
    }
    // 脏闭包 = 该格 + 传递依赖它的公式格
    const dirty = this.closure(key)
    const before = new Map<string, Scalar>()
    for (const k of dirty) before.set(k, this.cellAt(k)?.value ?? null)
    this.recompute([...dirty])
    // 收集值变了的格
    const out: DirtyCell[] = []
    for (const k of dirty) {
      const cur = this.cellAt(k)?.value ?? null
      if (!sameVal(before.get(k) ?? null, cur)) { const [s, r, c2] = k.split(':').map(Number); out.push({ sheet: s, row: r, col: c2, value: toCellValue(cur) }) }
    }
    // 改的格本身(字面值)也算变更
    if (!dirty.has(key)) { const cur = this.cellAt(key)?.value ?? null; out.push({ sheet, row, col, value: toCellValue(cur) }) }
    return out
  }

  getValue(sheet: number, row: number, col: number): CellValue {
    return toCellValue(this.sheets[sheet]?.cells.get(KEY(sheet, row, col))?.value ?? null)
  }

  destroy(): void {
    this.sheets = []
    this.nameToIndex.clear()
    this.dependents.clear()
  }

  // ---------------- 内部 ----------------
  private makeCell(sheet: number, _row: number, _col: number, content: string | number | boolean, _isEdit: boolean): EngCell {
    if (typeof content === 'string' && content[0] === '=') {
      try {
        const ast = parseFormula(content)
        return { formula: content, ast, value: null, prec: extractPrec(ast, sheet, this.nameToIndex) }
      } catch {
        return { formula: content, ast: null, value: ERR.name(), prec: [] } // 解析失败 → #NAME?
      }
    }
    return { value: content as Scalar }
  }

  private cellAt(key: string): EngCell | undefined {
    const i = key.indexOf(':')
    return this.sheets[Number(key.slice(0, i))]?.cells.get(key)
  }

  private addDeps(formulaKey: string, prec?: Prec[]): void {
    for (const p of prec ?? []) this.eachPrecCell(p, (ck) => { let set = this.dependents.get(ck); if (!set) this.dependents.set(ck, (set = new Set())); set.add(formulaKey) })
  }
  private removeDeps(formulaKey: string, prec: Prec[]): void {
    for (const p of prec) this.eachPrecCell(p, (ck) => { this.dependents.get(ck)?.delete(formulaKey) })
  }
  private eachPrecCell(p: Prec, fn: (key: string) => void): void {
    let r1 = p.r1, c1 = p.c1
    // 夹到上限,防整列/整行引用爆内存
    while ((r1 - p.r0 + 1) * (c1 - p.c0 + 1) > RANGE_CAP) { if (r1 - p.r0 > c1 - p.c0) r1--; else c1--; if (r1 < p.r0 || c1 < p.c0) break }
    for (let r = p.r0; r <= r1; r++) for (let c = p.c0; c <= c1; c++) fn(KEY(p.sheet, r, c))
  }

  /** 传递闭包:start + 所有(直接/间接)依赖它的公式格。 */
  private closure(start: string): Set<string> {
    const seen = new Set<string>([start])
    const stack = [start]
    while (stack.length) {
      const k = stack.pop()!
      const deps = this.dependents.get(k)
      if (deps) for (const d of deps) if (!seen.has(d)) { seen.add(d); stack.push(d) }
    }
    return seen
  }

  /** 递归求值一组格(memo 保证拓扑序;computing 栈检测循环 → #REF!)。 */
  private recompute(keys: string[]): void {
    const done = new Set<string>()
    const computing = new Set<string>()
    const evalCell = (key: string): Scalar => {
      const cell = this.cellAt(key)
      if (!cell) return null
      if (cell.ast === undefined) return cell.value // 字面值格
      if (done.has(key)) return cell.value
      if (computing.has(key)) return ERR.ref() // 循环引用
      if (cell.ast === null) { done.add(key); return cell.value } // 解析失败,固定错误
      computing.add(key)
      const sheetIdx = Number(key.slice(0, key.indexOf(':')))
      const resolver: RefResolver = {
        getCell: (sheetName, row, col) => {
          const si = sheetName == null ? sheetIdx : this.nameToIndex.get(sheetName)
          if (si == null) return ERR.ref()
          const pk = KEY(si, row, col)
          const pc = this.sheets[si]?.cells.get(pk)
          if (!pc) return null
          if (pc.ast !== undefined && !done.has(pk)) return evalCell(pk) // 脏前驱先算(memo)
          return pc.value
        },
      }
      const v = scalarResult(evalAst(cell.ast, resolver))
      computing.delete(key)
      done.add(key)
      cell.value = v
      return v
    }
    for (const k of keys) evalCell(k)
  }
}

// ---------------- 辅助 ----------------
function cellModelContent(cell: { type: string; formula?: string; raw: unknown }): string | number | boolean | null {
  if (cell.type === 'formula') { const f = cell.formula; if (!f) return null; return f[0] === '=' ? f : '=' + f }
  const raw = cell.raw
  if (typeof raw === 'number' || typeof raw === 'boolean' || typeof raw === 'string') return raw
  return null
}
function extractPrec(node: Node, curSheet: number, nameIdx: Map<string, number>): Prec[] {
  const out: Prec[] = []
  const walk = (n: Node): void => {
    switch (n.t) {
      case 'ref': { const s = n.sheet == null ? curSheet : nameIdx.get(n.sheet); if (s != null) out.push({ sheet: s, r0: n.row, c0: n.col, r1: n.row, c1: n.col }); break }
      case 'range': { const s = n.a.sheet == null ? curSheet : nameIdx.get(n.a.sheet); if (s != null) out.push({ sheet: s, r0: Math.min(n.a.row, n.b.row), c0: Math.min(n.a.col, n.b.col), r1: Math.max(n.a.row, n.b.row), c1: Math.max(n.a.col, n.b.col) }); break }
      case 'unary': walk(n.x); break
      case 'bin': walk(n.l); walk(n.r); break
      case 'func': n.args.forEach(walk); break
    }
  }
  walk(node)
  return out
}
function scalarResult(v: unknown): Scalar {
  if (v && typeof v === 'object' && '__m' in (v as object)) { const rows = (v as { rows: Scalar[][] }).rows; return rows.length && rows[0].length ? rows[0][0] : ERR.value() }
  return v as Scalar
}
function toCellValue(v: Scalar): CellValue {
  if (isErr(v)) return v.code
  return v
}
function sameVal(a: Scalar, b: Scalar): boolean {
  if (a instanceof FErr && b instanceof FErr) return a.code === b.code
  return a === b
}
