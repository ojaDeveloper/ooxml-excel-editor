/**
 * Vue 2 子入口 (`ooxml-excel-editor/vue2`, 2026-06-08, 1.3.0).
 *
 * 复用 `dist/core.js` 的同一份框架无关引擎 (ViewerController + 模型 + 渲染 + 编辑 + 导出),
 * 跟 Vue 3 / React 壳 ~100% 共享 core. 仅薄壳层使用 Vue 2.7+ Composition API.
 *
 * 用法:
 *   import ExcelViewer from 'ooxml-excel-editor/vue2'
 *   import 'ooxml-excel-editor/style.css'
 */
export { default as ExcelViewer } from './ExcelViewer'
export { default } from './ExcelViewer'

// 公共类型 / 工具 全部复用 core 已有出口
export * from '@/core/plugin'
export type {
  WorkbookModel,
  SheetModel,
  CellModel,
  CellStyle,
  CellStyleFn,
  CellStyleOverride,
  CellStyleCtx,
  MergeRange,
} from '@/core/model/types'
