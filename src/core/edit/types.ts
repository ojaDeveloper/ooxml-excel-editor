/**
 * 编辑配置(框架无关)。默认只读;开 editable 才进入编辑;可按格 / 按区域标只读。
 * 后续阶段(E1+)会在此扩展 recalc / formulaEngine 等;E0 只做"能不能编辑"的闸门。
 */
import type { CellModel, MergeRange } from '../model/types'
import type { FormulaEngineFactory } from '../formula/engine'

/**
 * 行/列维度目标 (Phase B, 2026-06-08) —— 用于尺寸 API (setColumnWidth / setRowHeight /
 * autoFitColumns / resetColumnWidth ...) 的参数. 3 种形状自动识别:
 *
 *   ┌────────────────────┬──────────────────────────┐
 *   │ 形状                │ 含义                      │
 *   ├────────────────────┼──────────────────────────┤
 *   │ `number`           │ 单个 index               │
 *   │ `number[]`         │ 多 index (允许不相邻)     │
 *   │ `{ from, to }`     │ 闭区间范围               │
 *   └────────────────────┴──────────────────────────┘
 */
export type DimTarget =
  | number
  | number[]
  | { from: number; to: number }

/**
 * 可编辑目标 —— 用于 `EditConfig.editableTargets` 的白名单元素;接受 4 种形状,
 * 自动识别(看带哪些字段):
 *
 *   ┌──────────────────────────────────────┬─────────────────────────────┐
 *   │ 形状                                  │ 含义                         │
 *   ├──────────────────────────────────────┼─────────────────────────────┤
 *   │ `{ row, col }`                       │ 单格(命中精确这一格)         │
 *   │ `{ row }`                            │ 整行(该行所有列)             │
 *   │ `{ col }`                            │ 整列(该列所有行)             │
 *   │ `{ top, left, bottom, right }`       │ 矩形区域(0-based 闭区间)     │
 *   └──────────────────────────────────────┴─────────────────────────────┘
 *
 * 多个 target 可以**重叠 / 不相邻**;命中**任一**就算可编辑。
 */
export type EditableTarget =
  | { row: number; col: number }
  | { row: number; col?: undefined }
  | { col: number; row?: undefined }
  | MergeRange

export interface EditConfig {
  /** 总开关:默认 false = 只读(行为与历史完全一致) */
  editable?: boolean
  /** 按格只读判定:返回 true = 该格只读。cell 为空格时传 null。pos 为 0-based 行列。 */
  cellReadOnly?: (cell: CellModel | null, pos: { row: number; col: number }) => boolean | void
  /** 只读区域(0-based 闭区间);命中即只读 */
  readOnlyRanges?: MergeRange[]
  /**
   * **可编辑白名单**(2026-06-08 新增) —— 设了就是"白名单模式":默认只读,
   * 只有命中**任一** target 的格才可编辑。**未设(undefined)= 默认全可编辑**
   * (老行为)。**显式传空数组 `[]`** = 全只读(没格在白名单)。
   *
   * 用例: 协同编辑、点检表单只填几格、模板里只让用户改"金额"列等.
   *
   * 与黑名单组合(优先级): `editable=false` → 全只读 ► 不在白名单 → 只读 ►
   * 命中 `readOnlyRanges` → 只读 ► `cellReadOnly` 返 true → 只读 ► 否则可编辑.
   */
  editableTargets?: EditableTarget | EditableTarget[]
  /**
   * **严格尺寸闸门**(Phase B, 2026-06-08) —— 默认 `false`:setColumnWidth / setRowHeight /
   * autoFit / resetDimensions 等尺寸 API 仅受 `editable` 全局闸门控制 (老行为, 简单).
   *
   * 设 `true` + 启用了 `editableTargets` 白名单 → 升级语义:**该列/行至少有 1 格在白名单内**
   * 才能改它的宽高; 否则拒绝 + emit permission-denied (reason='dimension').
   *
   * 跟"白名单未覆盖 = 完全只读"的严格语义一致, 防"用户改不了数据但能改列宽行高"。
   */
  strictDimensions?: boolean
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
