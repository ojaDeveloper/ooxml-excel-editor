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
  /** 提交编辑(走命令栈 + 事件);editor 自行决定何时调 */
  commit(value: EditorCommitValue): void
  /** 取消编辑(不改模型) */
  cancel(): void
}

/** 编辑器工厂:返回 DOM 元素,或 { el, destroy }(destroy 在卸载时清理) */
export type CellEditorFactory = (ctx: CellEditorContext) => HTMLElement | { el: HTMLElement; destroy?: () => void }

/** 编辑器解析器:按格决定用哪个工厂(无 → 该格无自定义编辑器) */
export type EditorResolver = (cell: CellModel | null, pos: { row: number; col: number }) => CellEditorFactory | void
