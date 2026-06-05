/**
 * 公式 A1 引用移位(框架无关)—— 增删行列后重写公式文本里的单元格引用,使其继续指向同一逻辑格。
 * 例:在第 5 行上方插入一行 → `=A5+B6` 变 `=A6+B7`;删除被引用的行 → 该引用变 `#REF!`。
 *
 * 处理:绝对/相对($)、跨表限定(`Sheet1!A1` / `'My Sheet'!A1`)、区域(`A1:B2`,删除时收缩,
 * 全落在删除段才 #REF!)。跳过:双引号字符串字面量、函数名(引用后紧跟 `(`)、被标识符粘连的裸引用。
 */
import { colIndexToLetters } from '../layout/grid-metrics'
import type { WorkbookModel } from '../model/types'
import type { StructOp } from '../model/structure'

export type ShiftAxis = 'row' | 'col'
export interface ShiftSpec {
  axis: ShiftAxis
  /** 0-based 模型索引:插入/删除发生的位置 */
  at: number
  count: number
  mode: 'insert' | 'delete'
  /** 结构操作发生在哪张表(名);仅指向它的引用才重写 */
  targetSheet: string
  /** 该公式所在表(名);无 `Sheet!` 限定的引用归属它 */
  formulaSheet: string
}

const DEAD = Number.NaN
function lettersToColIndex(s: string): number {
  let n = 0
  const up = s.toUpperCase()
  for (let i = 0; i < up.length; i++) n = n * 26 + (up.charCodeAt(i) - 64)
  return n - 1
}

/** 移位一个 0-based 索引;落在删除段返回 DEAD(NaN)。 */
function shiftIndex(idx: number, spec: ShiftSpec): number {
  if (spec.mode === 'insert') return idx >= spec.at ? idx + spec.count : idx
  if (idx >= spec.at + spec.count) return idx - spec.count // 删除段之后:上移
  if (idx >= spec.at) return DEAD // 删除段内
  return idx
}

interface Ref {
  colAbs: string
  col: string
  rowAbs: string
  row: string
}
/** 移位单个引用;clampForRange=true 时删除段端点夹到 at(区域收缩)而非死亡。返回 null = #REF!。 */
function shiftRef(r: Ref, spec: ShiftSpec, clampForRange: boolean): Ref | null {
  let colIdx = lettersToColIndex(r.col)
  let rowIdx = parseInt(r.row, 10) - 1
  if (spec.axis === 'row') {
    const v = shiftIndex(rowIdx, spec)
    if (Number.isNaN(v)) {
      if (!clampForRange) return null
      rowIdx = spec.at // 夹紧
    } else rowIdx = v
  } else {
    const v = shiftIndex(colIdx, spec)
    if (Number.isNaN(v)) {
      if (!clampForRange) return null
      colIdx = spec.at
    } else colIdx = v
  }
  if (rowIdx < 0 || colIdx < 0) return null
  return { colAbs: r.colAbs, col: colIndexToLetters(colIdx), rowAbs: r.rowAbs, row: String(rowIdx + 1) }
}
const fmt = (r: Ref) => `${r.colAbs}${r.col}${r.rowAbs}${r.row}`

// 可选 sheet 限定 + 单格 或 区域。组:1=sheet,2-5=首引用,6-9=次引用(区域)。
const REF_PART = '(\\$?)([A-Za-z]{1,3})(\\$?)(\\d+)'
const RANGE_RE = new RegExp(`(?:('[^']+'|[A-Za-z_\\u4e00-\\u9fff][\\w.\\u4e00-\\u9fff]*)!)?${REF_PART}(?::${REF_PART})?`, 'g')

function unquoteSheet(s: string): string {
  return s.startsWith("'") ? s.slice(1, -1).replace(/''/g, "'") : s
}

/** 重写一条公式串里所有指向 targetSheet 的引用以适配增删行列。前导 `=` 原样保留。 */
export function shiftFormulaRefs(formula: string, spec: ShiftSpec): string {
  // 拆出双引号字符串字面量(偶数段=代码,奇数段=字面量,不动)
  const parts = formula.split(/("(?:[^"]|"")*")/)
  for (let i = 0; i < parts.length; i += 2) {
    const seg = parts[i]
    parts[i] = seg.replace(RANGE_RE, (m, sheet, ca1, c1, ra1, r1, ca2, c2, ra2, r2, offset: number) => {
      const after = seg[offset + m.length]
      if (!c2 && after === '(') return m // 函数名,非引用
      const before = seg[offset - 1]
      if (!sheet && before && /[A-Za-z0-9_$.]/.test(before)) return m // 被标识符粘连的裸引用
      const refSheet = sheet ? unquoteSheet(sheet) : spec.formulaSheet
      if (refSheet.toLowerCase() !== spec.targetSheet.toLowerCase()) return m // 不指向目标表
      const prefix = sheet ? sheet + '!' : ''
      if (!c2) {
        const out = shiftRef({ colAbs: ca1, col: c1, rowAbs: ra1, row: r1 }, spec, false)
        return out ? prefix + fmt(out) : '#REF!' // 单格死亡 → 整体 #REF!
      }
      // 区域:端点删除则收缩;全死才 #REF!
      const a = shiftRef({ colAbs: ca1, col: c1, rowAbs: ra1, row: r1 }, spec, true)
      const b = shiftRef({ colAbs: ca2, col: c2, rowAbs: ra2, row: r2 }, spec, true)
      const aDead = !shiftRef({ colAbs: ca1, col: c1, rowAbs: ra1, row: r1 }, spec, false)
      const bDead = !shiftRef({ colAbs: ca2, col: c2, rowAbs: ra2, row: r2 }, spec, false)
      if (aDead && bDead) return prefix + '#REF!'
      return prefix + fmt(a!) + ':' + fmt(b!)
    })
  }
  return parts.join('')
}

/**
 * 增删行列后,重写**全簿**所有公式格里指向 targetSheet 的引用(含跨表 `Sheet!A1`)。
 * 就地改 cell.formula;撤销由结构命令的整簿快照负责。
 */
export function rewriteWorkbookFormulas(wb: WorkbookModel, targetSheetIdx: number, op: StructOp, at: number, count: number): void {
  const target = wb.sheets[targetSheetIdx]
  if (!target) return
  const axis: ShiftAxis = op === 'insert-rows' || op === 'delete-rows' ? 'row' : 'col'
  const mode: 'insert' | 'delete' = op.startsWith('insert') ? 'insert' : 'delete'
  for (const sheet of wb.sheets) {
    for (const cell of sheet.cells.values()) {
      if (cell.type === 'formula' && cell.formula) {
        cell.formula = shiftFormulaRefs(cell.formula, { axis, at, count, mode, targetSheet: target.name, formulaSheet: sheet.name })
      }
    }
  }
}
