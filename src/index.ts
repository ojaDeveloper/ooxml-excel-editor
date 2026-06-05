/**
 * 组件库入口。其它项目这样用:
 *   import { ExcelViewer } from 'ooxml-excel-editor'
 *   import 'ooxml-excel-editor/style.css'
 * 或全局注册:
 *   import OoxmlExcelPreview from 'ooxml-excel-editor'
 *   app.use(OoxmlExcelPreview)
 */
import type { App, Plugin } from 'vue'
import ExcelViewer from './components/ExcelViewer.vue'

export { ExcelViewer }

// 程序化解析能力(不渲染，只要模型时用)
export { parseWorkbook } from './core/parser'
export { loadArrayBuffer } from './core/loader'
export type { ExcelSource } from './core/loader'

// 数据读取 API(配 parseWorkbook 独立用;组件 ref 也有同名方法)
export {
  cellDisplayText,
  getCell,
  getCellValue,
  getCellStyle,
  getCellText,
  getSheetData,
  getRangeData,
  sheetToJSON,
  getWorkbookJSON,
} from './core/model/data-access'
export type { CellValue, ReadOptions, SheetToJSONOptions } from './core/model/data-access'
// 顺带导出消费者常用的底层工具
export { formatValue } from './core/format/number-format'
export { cellKey } from './core/model/types'

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
  ShapeSpec,
  Sparkline,
  CssColor,
  TransformModelFn,
  CellStyleFn,
  CellStyleOverride,
} from './core/model/types'

// 扩展: 主题
export type { ViewerTheme } from './core/render/theme'
export { DEFAULT_THEME, mergeTheme } from './core/render/theme'

// 扩展: 插件
export { definePlugin } from './core/plugin'
export type {
  ExcelPlugin,
  ExcelPluginContext,
  ViewerApi,
  OverlayContext,
  PluginEvent,
  Rect,
  ToolbarItem,
} from './core/plugin'

// 扩展: 导出 / 打印(选项类型 + beforeRenderPage 钩子)。
// 命令式方法(exportImage/downloadImage/exportPdf/downloadPdf/print)在组件 ref(ViewerApi)上。
export type {
  ExportTarget,
  RenderExportOptions,
  ImageExportOptions,
  PdfExportOptions,
  PrintOptions,
  PageSetup,
  PageFormat,
  Orientation,
  Margins,
  PdfPageContext,
  BeforeRenderPage,
} from './core/export/types'
// 低层导出工具(自定义编排/独立使用时可选)
export { canvasToBlob, canvasToDataURL, downloadBlob } from './core/export/raster'

// 作为 Vue 插件全局注册
const plugin: Plugin = {
  install(app: App) {
    app.component('ExcelViewer', ExcelViewer)
  },
}

export default plugin
