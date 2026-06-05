/**
 * 框架无关 core 入口(供 'ooxml-excel-preview/core' 及 Vue / React 壳复用)。
 * 这里只暴露引擎 + 解析 + 数据 + 类型,零框架依赖 —— Vue/React 各自的壳在上层包。
 */

// ---- 解析 / 加载 / 收尾 ----
export { parseWorkbook } from './parser'
export { loadArrayBuffer } from './loader'
export type { ExcelSource } from './loader'
export { detectFormat, finalizeImages, friendlyError, revokeImages } from './finalize'
export type { ParseProgress } from './progress'

// ---- 中间数据模型 + 读数据 API ----
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
} from './model/data-access'
export type { CellValue, ReadOptions, SheetToJSONOptions } from './model/data-access'
export { cellKey } from './model/types'
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
} from './model/types'

// ---- 渲染引擎 / 控制器 / 叠加层 ----
export { CanvasRenderer } from './render/canvas-renderer'
export type { ViewState, RendererOptions } from './render/canvas-renderer'
export { ViewerController } from './viewer/controller'
export type {
  Cell,
  TooltipState,
  FindState,
  FilterPopupState,
  ViewerControllerEls,
  ViewerControllerHooks,
} from './viewer/controller'
export { OverlayManager } from './viewer/overlay-manager'
export type { OverlayQuads } from './viewer/overlay-manager'
export { PluginOverlayHost } from './viewer/plugin-overlay'

// ---- 编辑(配置 + 只读判定 + 写数据层 + 命令栈 + 快照;E0/E1) ----
export { resolveEditable } from './edit/permissions'
export type { EditConfig, EditPermission } from './edit/types'
export { EditController } from './edit/edit-controller'
export type {
  EditControllerHost,
  EditEventName,
  EditSource,
  CellChangePayload,
  DimChangePayload,
  DirtyChangePayload,
} from './edit/edit-controller'
export { isDimCommand } from './edit/commands'
export type { EditCommand, CellPos, DimAxis } from './edit/commands'
export { buildCellSnapshot, cloneCell } from './model/snapshot'
export type { CellSnapshot } from './model/snapshot'
export {
  setCellValue,
  clearCell,
  setRangeValues,
  restoreCell,
  internStyle,
  setColumnWidth,
  setRowHeight,
  restoreDimension,
} from './model/mutations'
export { cloneWorkbook, restoreWorkbookInto } from './model/clone'
export { CellEditorHost } from './edit/editor-host'
export type { CellEditorContext, CellEditorFactory, EditorResolver, EditorCommitValue } from './edit/editor-context'

// ---- 插件 / 扩展点(框架无关) ----
export { definePlugin } from './plugin'
export type {
  ExcelPlugin,
  ExcelPluginContext,
  OverlayContext,
  OverlayNode,
  PluginEvent,
  Rect,
  ToolbarItem,
  ViewerApi,
} from './plugin'

// ---- 布局 / 格式 ----
export { GridMetrics, colIndexToLetters } from './layout/grid-metrics'
export { formatValue } from './format/number-format'

// ---- 主题 ----
export { DEFAULT_THEME, mergeTheme } from './render/theme'
export type { ViewerTheme } from './render/theme'

// ---- 导出 / 打印 ----
export { WorkbookExporter } from './export/exporter'
export type { ExporterHost } from './export/exporter'
export { canvasToBlob, canvasToDataURL, downloadBlob } from './export/raster'
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
} from './export/types'
