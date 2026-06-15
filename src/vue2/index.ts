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

// 框架无关 core 全部公共出口(parseWorkbook / 数据读取 / 模型类型 / definePlugin / 公式引擎工厂 / 导出类型…)
// 跟主入口、`/core`、`/react` 同源 —— 不再各自维护清单致漂移。
export * from '@/core'
