/**
 * 组件库入口。其它项目这样用:
 *   import { ExcelViewer } from 'ooxml-excel-preview'
 *   import 'ooxml-excel-preview/style.css'
 * 或全局注册:
 *   import OoxmlExcelPreview from 'ooxml-excel-preview'
 *   app.use(OoxmlExcelPreview)
 */
import type { App, Plugin } from 'vue'
import ExcelViewer from './components/ExcelViewer.vue'

export { ExcelViewer }

// 程序化解析能力(不渲染，只要模型时用)
export { parseWorkbook } from './core/parser'
export { loadArrayBuffer } from './core/loader'
export type { ExcelSource } from './core/loader'

// 中间数据模型类型，供宿主项目做二次处理/类型标注
export type {
  WorkbookModel,
  SheetModel,
  CellModel,
  CellStyle,
  MergeRange,
  ConditionalRule,
  ChartSpec,
  ImageAnchor,
  CssColor,
} from './core/model/types'

// 作为 Vue 插件全局注册
const plugin: Plugin = {
  install(app: App) {
    app.component('ExcelViewer', ExcelViewer)
  },
}

export default plugin
