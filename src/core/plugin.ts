/**
 * 插件接口 —— 把"扩展点"(主题/数据钩子/渲染钩子/事件/overlay/命令式 API)打包分发。
 * 用法:
 *   const myPlugin = definePlugin({
 *     name: 'highlight-negatives',
 *     cellStyle: (c) => typeof c.raw === 'number' && c.raw < 0 ? { font: { color: '#d00' } } : undefined,
 *     events: { 'cell-click': (p) => console.log(p) },
 *     overlay: ({ rectOf }) => rectOf(0,0) ? h('div', ...) : null,
 *     setup: ({ viewer, on }) => { on('selection-change', ...); return () => {} },
 *   })
 *   <ExcelViewer :plugins="[myPlugin]" />
 */
import type { VNodeChild } from 'vue'
import type { CellStyleFn, MergeRange, TransformModelFn, WorkbookModel } from './model/types'
import type { ViewerTheme } from './render/theme'
import type { ExcelSource } from './loader'
import type { ImageExportOptions, PdfExportOptions, PrintOptions } from './export/types'

export interface Rect {
  x: number
  y: number
  w: number
  h: number
}

export type PluginEvent =
  | 'cell-click'
  | 'cell-dblclick'
  | 'selection-change'
  | 'sheet-change'
  | 'hyperlink-click'

/** 命令式 API(组件 ref 与插件 ctx 共用) */
export interface ViewerApi {
  load(src: ExcelSource): void
  getWorkbook(): WorkbookModel | null
  getActiveSheet(): number
  setActiveSheet(index: number): void
  getSelection(): MergeRange | null
  setSelection(range: MergeRange): void
  rectOf(row: number, col: number): Rect | null
  rectOfRange(range: MergeRange): Rect | null
  redraw(): void
  /** 导出当前/指定表为图片 Blob(默认 png) */
  exportImage(opts?: ImageExportOptions): Promise<Blob>
  /** 导出为图片并触发下载 */
  downloadImage(opts?: ImageExportOptions): Promise<void>
  /** 导出为 PDF Blob(需可选依赖 jspdf);beforeRenderPage 可画页眉/页脚/水印 */
  exportPdf(opts?: PdfExportOptions): Promise<Blob>
  /** 导出 PDF 并触发下载 */
  downloadPdf(opts?: PdfExportOptions): Promise<void>
  /** 打开系统打印(可在对话框另存为 PDF) */
  print(opts?: PrintOptions): Promise<void>
}

/** overlay 渲染上下文(随滚动/缩放,tick 变即重渲) */
export interface OverlayContext {
  rectOf(row: number, col: number): Rect | null
  rectOfRange(range: MergeRange): Rect | null
  tick: number
  workbook: WorkbookModel | null
}

export interface ExcelPluginContext {
  viewer: ViewerApi
  /** 订阅交互事件 */
  on(event: PluginEvent, handler: (payload: any) => void): void
  /** 主动重绘 */
  redraw(): void
}

export interface ExcelPlugin {
  name: string
  /** 外观主题覆盖(多插件按数组顺序合并,组件 :theme 最后覆盖) */
  theme?: Partial<ViewerTheme>
  /** 数据钩子: 解析后改模型(多插件 + 组件 prop 链式应用) */
  transformModel?: TransformModelFn
  /** 渲染钩子: 按单元格覆盖样式(多插件 + 组件 prop 合并) */
  cellStyle?: CellStyleFn
  /** 交互事件处理(简单写法;复杂用 setup 的 on) */
  events?: Partial<Record<PluginEvent, (payload: any) => void>>
  /** 在网格上叠加 UI(返回 VNode);随 tick 重渲 */
  overlay?: (ctx: OverlayContext) => VNodeChild
  /** 高级: 拿命令式 API、订阅事件;返回可选清理函数 */
  setup?: (ctx: ExcelPluginContext) => void | (() => void)
}

/** 定义插件(仅作类型推断,原样返回) */
export function definePlugin(plugin: ExcelPlugin): ExcelPlugin {
  return plugin
}
