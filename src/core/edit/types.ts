/**
 * 编辑配置(框架无关)。默认只读;开 editable 才进入编辑;可按格 / 按区域标只读。
 * 后续阶段(E1+)会在此扩展 recalc / formulaEngine 等;E0 只做"能不能编辑"的闸门。
 */
import type { CellModel, MergeRange } from '../model/types'

export interface EditConfig {
  /** 总开关:默认 false = 只读(行为与历史完全一致) */
  editable?: boolean
  /** 按格只读判定:返回 true = 该格只读。cell 为空格时传 null。pos 为 0-based 行列。 */
  cellReadOnly?: (cell: CellModel | null, pos: { row: number; col: number }) => boolean | void
  /** 只读区域(0-based 闭区间);命中即只读 */
  readOnlyRanges?: MergeRange[]
}

/** 单元格编辑权限 */
export type EditPermission = 'editable' | 'readonly'
