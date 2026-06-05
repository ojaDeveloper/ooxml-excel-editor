/**
 * 插件接口 —— 把"扩展点"(主题/数据钩子/渲染钩子/事件/overlay/命令式 API)打包分发。
 * 用法:
 *   const myPlugin = definePlugin({
 *     name: 'highlight-negatives',
 *     cellStyle: (c) => typeof c.raw === 'number' && c.raw < 0 ? { font: { color: '#d00' } } : undefined,
 *     events: { 'cell-click': (p) => console.log(p) },
 *     overlay: ({ rectOf }) => { const r = rectOf(0,0); if (!r) return null
 *       const el = document.createElement('div'); el.textContent = '⚑'
 *       Object.assign(el.style, { position:'absolute', left:r.x+'px', top:r.y+'px' }); return el },
 *     setup: ({ viewer, on }) => { on('selection-change', ...); return () => {} },
 *   })
 *   <ExcelViewer :plugins="[myPlugin]" />   // Vue
 *   <ExcelViewer plugins={[myPlugin]} />    // React —— 同一插件,两框架通用
 */
import type { CellStyleFn, CellStyleOverride, ImageAnchor, MergeRange, TransformModelFn, WorkbookModel } from './model/types'
import type { CellValue, ReadOptions, SheetToJSONOptions } from './model/data-access'
import type { CellSnapshot } from './model/snapshot'
import type { EditorResolver } from './edit/editor-context'
import type { ViewerTheme } from './render/theme'
import type { ExcelSource } from './loader'
import type { ImageExportOptions, PdfExportOptions, PrintOptions } from './export/types'
import type { XlsxExportOptions } from './export/xlsx-writer'

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
  | 'cell-change'
  | 'edit-start'
  | 'edit-commit'
  | 'dim-change'
  | 'dirty-change'
  | 'image-change'
  | 'struct-change'

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
  /** 该格当前是否可编辑(综合 editable + readOnlyRanges + cellReadOnly) */
  isCellEditable(row: number, col: number): boolean
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
  // ---- 数据导出(E8;一份数据层 → xlsx/json/csv) ----
  /** 整簿 → .xlsx Blob(从编辑后模型重建,所见即所得;需可选依赖 exceljs) */
  exportXlsx(opts?: XlsxExportOptions): Promise<Blob>
  /** 导出 .xlsx 并触发下载 */
  downloadXlsx(opts?: XlsxExportOptions): Promise<void>
  /** 整簿 → JSON 文本(各表首行作 key) */
  exportJson(opts?: SheetToJSONOptions): string
  /** 导出 JSON 并触发下载 */
  downloadJson(opts?: SheetToJSONOptions): void
  /** 一张表 → CSV 文本(默认活动表、格式化显示值) */
  exportCsv(opts?: { target?: number; format?: boolean }): string
  /** 导出 CSV 并触发下载(带 UTF-8 BOM) */
  downloadCsv(opts?: { target?: number; format?: boolean }): void
  // ---- 数据读取(自动用当前 workbook 的 date1904;sheetIndex 缺省=当前活动表) ----
  /** 单元格原始值 */
  getCellValue(row: number, col: number, sheetIndex?: number): CellValue
  /** 单元格格式化显示文本 */
  getCellText(row: number, col: number, sheetIndex?: number): string
  /** 整表二维数组(format 默认 true=显示文本) */
  getSheetData(opts?: ReadOptions, sheetIndex?: number): CellValue[][]
  /** 整表对象数组(首行作表头) */
  getSheetJSON(opts?: SheetToJSONOptions, sheetIndex?: number): Record<string, CellValue>[]
  /** 区域二维数组 */
  getRangeData(range: MergeRange, opts?: ReadOptions, sheetIndex?: number): CellValue[][]
  // ---- 编辑(E1;需 editable 开启,只读格不生效) ----
  /** 编辑单格;返回是否生效 */
  editCell(row: number, col: number, value: CellValue): boolean
  /** 区域批量设值(2D,左上对齐 range.top/left);跳过只读格 */
  editRange(range: MergeRange, values: CellValue[][]): boolean
  /** 清空区域(跳过只读) */
  clearRange(range: MergeRange): boolean
  undo(): void
  redo(): void
  canUndo(): boolean
  canRedo(): boolean
  /** 当前正在编辑的格(无则 null) */
  getEditingCell(): { row: number; col: number } | null
  /** 查询任意格的完整快照(底层结构 + raw/computed/text/style) */
  getCellSnapshot(row: number, col: number): CellSnapshot | null
  /** 进入编辑(需有 editor 工厂 + 可编辑);返回是否进入 */
  beginEdit(row: number, col: number): boolean
  /** 取消当前编辑(不改模型) */
  cancelEdit(): void
  /** 当前是否有活动编辑器 */
  isEditing(): boolean
  /** 给区域套样式覆盖(E5;粗体/对齐/填充等);editable 时入命令栈(可撤销 + 发 cell-change + 记脏) */
  setStyle(range: MergeRange, patch: CellStyleOverride): boolean
  /** 合并区域(G1;清空被覆盖格,只留左上锚点);editable 时入命令栈 */
  mergeCells(range: MergeRange): boolean
  /** 拆分区域内的合并(G1);editable 时入命令栈 */
  unmergeCells(range: MergeRange): boolean
  /** 读当前表全部图片锚点(克隆;E6) */
  getImages(): ImageAnchor[]
  /** 加一张图(无 src 但有 bytes+mime 时自动生成 blob url);返回插入索引 */
  addImage(anchor: ImageAnchor): number
  /** 删一张图 */
  removeImage(index: number): boolean
  /** 移动图片(屏幕像素增量);editable 时入命令栈 + 发 image-change */
  moveImage(index: number, dxPx: number, dyPx: number): boolean
  /** 缩放图片(目标屏幕像素宽高);editable 时入命令栈 + 发 image-change */
  resizeImage(index: number, widthPx: number, heightPx: number): boolean
  /** 在 at 处插入 count 行(E7);editable 时入命令栈 + 发 struct-change */
  insertRows(at: number, count?: number): boolean
  /** 删除 [at, at+count) 行(与合并相交则相交合并被移除) */
  deleteRows(at: number, count?: number): boolean
  /** 在 at 处插入 count 列 */
  insertCols(at: number, count?: number): boolean
  /** 删除 [at, at+count) 列 */
  deleteCols(at: number, count?: number): boolean
  /** 程序化设列宽(px,模型单位);editable 时入命令栈(可撤销 + 发 dim-change + 记脏) */
  setColumnWidth(col: number, width: number): boolean
  /** 程序化设行高(px,模型单位);editable 时入命令栈 */
  setRowHeight(row: number, height: number): boolean
  /** 公式引擎是否已就绪(recalc 开启 + 异步 warm 完成);未开重算恒 false */
  isRecalcReady(): boolean
  /** 当前是否有未保存修改(自加载/还原以来发生过编辑或 resize) */
  isDirty(): boolean
  /** 放弃全部修改,还原到刚加载的原件;返回是否还原 */
  resetToOriginal(): boolean
}

/** overlay 渲染上下文(随滚动/缩放,tick 变即重渲) */
export interface OverlayContext {
  rectOf(row: number, col: number): Rect | null
  rectOfRange(range: MergeRange): Rect | null
  tick: number
  workbook: WorkbookModel | null
}

/** overlay 钩子返回值:框架无关的 DOM 节点(单个 / 多个 / 无)。Vue 与 React 壳都直接挂载。 */
export type OverlayNode = HTMLElement | HTMLElement[] | null

export interface ExcelPluginContext {
  viewer: ViewerApi
  /** 订阅交互事件 */
  on(event: PluginEvent, handler: (payload: any) => void): void
  /** 主动重绘 */
  redraw(): void
}

/** 操作工具栏的一个按钮(插件贡献 / 组件 :toolbar 配置 / 命令式自定义共用) */
export interface ToolbarItem {
  /** 唯一 id(内置: 'find'|'filter'|'sort'|'export'|'zoom'|'copy'|'freeze'|'clear-filter';自定义任意) */
  id: string
  /** 类型: 'button'(默认) | 'separator'(分隔线,其余字段忽略) */
  type?: 'button' | 'separator'
  /** 图标(emoji / 字形);跨平台一致建议用 iconSvg */
  icon?: string
  /** 内联 SVG(优先于 icon);建议 24×24 viewBox、stroke=currentColor */
  iconSvg?: string
  /** 文字 */
  label?: string
  /** 悬停提示 */
  title?: string
  /** 点击回调(拿到命令式 API);有 items 时点击改为展开下拉 */
  onClick?: (viewer: ViewerApi) => void
  /** 是否高亮(激活态),如开关类按钮 */
  active?: (viewer: ViewerApi) => boolean
  /** 是否禁用(置灰不可点) */
  disabled?: (viewer: ViewerApi) => boolean
  /** 下拉子菜单项;有则本项为下拉按钮(点击展开) */
  items?: ToolbarItem[]
}

export interface ExcelPlugin {
  name: string
  /** 外观主题覆盖(多插件按数组顺序合并,组件 :theme 最后覆盖) */
  theme?: Partial<ViewerTheme>
  /** 贡献操作工具栏按钮(opt-in: 插件加载即出现) */
  toolbar?: ToolbarItem[]
  /** 数据钩子: 解析后改模型(多插件 + 组件 prop 链式应用) */
  transformModel?: TransformModelFn
  /** 渲染钩子: 按单元格覆盖样式(多插件 + 组件 prop 合并) */
  cellStyle?: CellStyleFn
  /** 交互事件处理(简单写法;复杂用 setup 的 on) */
  events?: Partial<Record<PluginEvent, (payload: any) => void>>
  /** 在网格上叠加 UI(返回 DOM 节点,框架无关);随 tick 重渲。用 ctx.rectOf 定位单元格。 */
  overlay?: (ctx: OverlayContext) => OverlayNode
  /** 按格自定义编辑控件(返回工厂;多插件数组序首个非空胜,组件 editor prop 覆盖);需 editable 开启。 */
  editor?: EditorResolver
  /** 高级: 拿命令式 API、订阅事件;返回可选清理函数 */
  setup?: (ctx: ExcelPluginContext) => void | (() => void)
}

/** 定义插件(仅作类型推断,原样返回) */
export function definePlugin(plugin: ExcelPlugin): ExcelPlugin {
  return plugin
}
