/**
 * 编辑配置(框架无关)。默认只读;开 editable 才进入编辑;可按格 / 按区域标只读。
 * 后续阶段(E1+)会在此扩展 recalc / formulaEngine 等;E0 只做"能不能编辑"的闸门。
 */
import type { CellModel, MergeRange } from '../model/types'
import type { FormulaEngineFactory } from '../formula/engine'

export interface EditConfig {
  /** 总开关:默认 false = 只读(行为与历史完全一致) */
  editable?: boolean
  /** 按格只读判定:返回 true = 该格只读。cell 为空格时传 null。pos 为 0-based 行列。 */
  cellReadOnly?: (cell: CellModel | null, pos: { row: number; col: number }) => boolean | void
  /** 只读区域(0-based 闭区间);命中即只读 */
  readOnlyRanges?: MergeRange[]
  /**
   * 公式重算(E4):默认 false = 沿用 Excel 缓存值(只读/无公式路径零成本)。
   * 开启后,编辑公式格或被公式引用的格 → 依赖格自动重算并逐个发 cell-change。
   * 默认引擎 HyperFormula(GPL-3.0/商业双授权,可选 peer);需 `npm i hyperformula`。
   */
  recalc?: boolean
  /** 自定义/自研公式引擎工厂(可换引擎);不给则用默认 HyperFormula 适配器。 */
  formulaEngine?: FormulaEngineFactory
}

/** 单元格编辑权限 */
export type EditPermission = 'editable' | 'readonly'
