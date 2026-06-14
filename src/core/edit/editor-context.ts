/**
 * 单元格编辑器扩展点(框架无关)。用方/插件提供一个工厂,返回任意 DOM 当编辑控件
 * (下拉 / 日期选择器 / 图片选择器 / 带按钮的面板…),用 ctx.commit/cancel 收口。
 * 这是要求 2(自定义编辑)的承重钩子;E3 的内置文本编辑器就是它的默认实现。
 */
import type { SheetModel, WorkbookModel, CellModel, CellStyleOverride } from '../model/types'
import type { CellValue } from '../model/data-access'
import type { CellSnapshot } from '../model/snapshot'
import type { EditPermission } from './types'

export type Rect = { x: number; y: number; w: number; h: number }

/** commit 既可只给值,也可同时给样式(样式 E5 起生效) */
export type EditorCommitValue = CellValue | { value: CellValue; style?: CellStyleOverride }

export interface CellEditorContext {
  /** 进入编辑时的完整前态快照 */
  snapshot: CellSnapshot
  /** 单元格当前屏幕矩形(render-area 相对) */
  rect: Rect
  sheet: SheetModel
  workbook: WorkbookModel
  permission: EditPermission
  /** 进入编辑的初始文本(打字进入时为该字符;否则 undefined → editor 用 snapshot.text) */
  initialText?: string
  /**
   * 提交编辑(走命令栈 + 事件);move 指示提交后活动格移动方向。
   * 返回 false = 提交被拒(如数据验证拦截),编辑器应保持打开让用户改正;void/true = 已提交。
   */
  commit(value: EditorCommitValue, move?: 'down' | 'right'): boolean | void
  /** 取消编辑(不改模型) */
  cancel(): void
  /**
   * 请求 host 重新 position 编辑器 (Phase 1 长文本撑高用, 2026-06-08).
   * 编辑器在内容变化(input 事件)后调一次, host 会重新读 `getDesiredHeight()` 并撑高.
   * 不实现 `getDesiredHeight` 的编辑器调它无效(仍按 cell rect 尺寸).
   */
  reposition?(): void
}

/**
 * 编辑器工厂的返回类型:可以是裸 HTMLElement, 或带钩子的对象.
 * - `destroy`: 卸载时清理(可选)
 * - `getDesiredHeight`: Phase 1 长文本撑高用. host 在 position 时调, 传入当前列宽 (px),
 *   编辑器返回内容期望的总高度 (px). host 取 max(原 cell h, desired) 作为最终高度.
 *   不实现 = 高度锁定为单元格高度(老行为).
 */
export interface CellEditorReturn {
  el: HTMLElement
  destroy?: () => void
  getDesiredHeight?(widthPx: number): number
}

/** 编辑器工厂:返回 DOM 元素,或 { el, destroy?, getDesiredHeight? }(destroy 在卸载时清理) */
export type CellEditorFactory = (ctx: CellEditorContext) => HTMLElement | CellEditorReturn

/** 编辑器解析器:按格决定用哪个工厂(无 → 该格无自定义编辑器) */
export type EditorResolver = (cell: CellModel | null, pos: { row: number; col: number }) => CellEditorFactory | void
