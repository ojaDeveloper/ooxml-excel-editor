/**
 * React 入口。其它 React 项目这样用:
 *   import { ExcelViewer } from 'ooxml-excel-editor/react'
 *   import 'ooxml-excel-editor/react/style.css'   // (打包后)
 *
 * 与 Vue 版共用同一套 core 引擎(ViewerController + 解析/渲染/导出)。
 */
export { ExcelViewer } from './ExcelViewer'
export type { ExcelViewerProps, ExcelViewerHandle } from './ExcelViewer'
export { useExcelDocument } from './use-excel-document'
export type { ExcelDocument } from './use-excel-document'

// 框架无关 core 全部公共出口(parseWorkbook / 数据读取 / 模型类型 / definePlugin / 公式引擎工厂 / 导出类型…)
// 跟主入口、`/core`、`/vue2` 同源,React 用方拿到同一套 —— 解决此前 React 入口太薄、core 能力够不着的问题。
export * from '../core'
