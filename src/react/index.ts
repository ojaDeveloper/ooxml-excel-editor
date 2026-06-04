/**
 * React 入口。其它 React 项目这样用:
 *   import { ExcelViewer } from 'ooxml-excel-preview/react'
 *   import 'ooxml-excel-preview/react/style.css'   // (打包后)
 *
 * 与 Vue 版共用同一套 core 引擎(ViewerController + 解析/渲染/导出)。
 */
export { ExcelViewer } from './ExcelViewer'
export type { ExcelViewerProps, ExcelViewerHandle } from './ExcelViewer'
export { useExcelDocument } from './use-excel-document'
export type { ExcelDocument } from './use-excel-document'
