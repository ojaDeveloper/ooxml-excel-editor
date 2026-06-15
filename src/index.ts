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

// 框架无关 core 的全部公共出口(解析/数据/类型/主题/插件/导出/**公式引擎工厂**/编辑/渲染…)
// 一次性 re-export,跟 `/core`、`/react`、`/vue2` 入口完全同源 —— 不再各自维护清单致漂移。
// 含:parseWorkbook · loadArrayBuffer · 数据读取 API · 模型类型 · formatValue · cellKey · definePlugin ·
//     主题(DEFAULT_THEME/mergeTheme)· 导出类型 · canvasToBlob… · ViewerApi 类型 ·
//     公式引擎(builtinFormulaEngineFactory 默认 / hyperFormulaEngineFactory / FUNCTION_NAMES / FormulaEngine 类型)。
export * from './core'

// 作为 Vue 插件全局注册
const plugin: Plugin = {
  install(app: App) {
    app.component('ExcelViewer', ExcelViewer)
  },
}

export default plugin
