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
import type { CellStyleFn, CellStyleOverride, ConditionalRule, ImageAnchor, MergeRange, PivotTableLayout, TransformModelFn, WorkbookModel } from './model/types'
import type { CellValue, ReadOptions, SheetToJSONOptions } from './model/data-access'
import type { CellSnapshot } from './model/snapshot'
import type { CellInspection } from './model/inspect'
import type { MenuItem } from './edit/context-menu'
import type { EditorResolver } from './edit/editor-context'
import type { EditableTarget } from './edit/types'
import type { PasteBehavior } from './edit/paste-behavior'
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

export type PivotOutput = { kind: 'current-sheet'; cell: string } | { kind: 'new-sheet' }
export interface CreatePivotTableOptions {
  /** 源数据区域,第一行作为字段名。缺省时使用当前选区。 */
  sourceRange?: MergeRange
  /** 源数据所在 sheet index。缺省为当前活动表。 */
  sourceSheetIndex?: number
  /** 输出位置。缺省为当前表源区域右侧空两列。 */
  output?: PivotOutput
  /** 透视表布局。缺省:第一个文本字段为行字段,第一个数值字段为值字段。 */
  layout?: Partial<PivotTableLayout>
  /** 是否打开右侧字段面板。缺省 false;工具栏入口会打开。 */
  showPanel?: boolean
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
  | 'permission-denied'

/**
 * 权限拒绝事件 payload (Phase A, 2026-06-08):mutation 因 editable / editableTargets /
 * readOnlyRanges / cellReadOnly 被阻挡时, 一次操作结束**统一**发一次. 默认行为仍是
 * **静默跳过**(跟 editRange 一致), 此事件**只通知**消费方做 toast / 高亮, 不阻断流程.
 *
 * 一次操作只 emit 一次 (避免 N 张图 spam N 次).
 */
export interface PermissionDeniedPayload {
  /** 触发的操作类型('copy' = 复制图片超字节预算,已降级为无图复制,非真正"拒绝") */
  reason: 'paste' | 'merge' | 'unmerge' | 'image-place' | 'image-convert' | 'dimension' | 'copy' | 'other'
  /** 被拒的目标格 (粘贴 / 合并 / 图片转换 等场景下的具体位置;'dimension' 时可空) */
  cells: Array<{ row: number; col: number }>
  /** 'dimension' 场景下被拒的列 / 行 index 列表 */
  dims?: { axis: 'col' | 'row'; indices: number[] }
  /** 给消费方的可读说明 */
  message?: string
}

/** 命令式 API(组件 ref 与插件 ctx 共用) */
export interface ViewerApi {
  load(src: ExcelSource): void
  getWorkbook(): WorkbookModel | null
  getActiveSheet(): number
  setActiveSheet(index: number): void
  getSelection(): MergeRange | null
  setSelection(range: MergeRange): void
  /** 滚动到指定单元格;select=true 时同步选中目标格。 */
  scrollToCell(row: number, col: number, opts?: { select?: boolean }): boolean
  rectOf(row: number, col: number): Rect | null
  rectOfRange(range: MergeRange): Rect | null
  redraw(): void
  /** 该格当前是否可编辑(综合 editable + editableTargets 白名单 + readOnlyRanges + cellReadOnly) */
  isCellEditable(row: number, col: number): boolean
  /**
   * **运行时**改可编辑白名单 (2026-06-08 新增). 接受 4 种 target 形状:
   * `{row,col}` 单格 / `{row}` 整行 / `{col}` 整列 / `MergeRange` 矩形;单值或数组都支持,允许**不相邻**.
   * 传 `undefined` = 关闭白名单(默认全可编辑);`[]` = 全只读;命中**任一** target → 可编辑.
   * 立即生效, 不动 `:editableTargets` prop.
   */
  setEditableTargets(targets: EditableTarget | EditableTarget[] | undefined): void
  /** 当前生效的可编辑白名单. `undefined` 表示未启用白名单. */
  getEditableTargets(): EditableTarget | EditableTarget[] | undefined
  /** 按活动单元格所在列排序;未开启自动筛选时会先按选区/已用区建立筛选范围。 */
  sortActiveColumn(dir: 'asc' | 'desc'): boolean
  /** 通过 API 直接创建静态透视表,不依赖当前页面选区或对话框。需开启 `pivotTable` + `editable` 配置(默认均关)。 */
  createPivotTable(opts: CreatePivotTableOptions): boolean
  /** 基于当前选区创建静态透视汇总表;未传 opts 时使用默认布局并输出到右侧。需 `pivotTable` + `editable`。 */
  createPivotTableFromSelection(opts?: {
    rowFieldIndex?: number
    valueFieldIndex?: number
    output?: PivotOutput
  }): boolean
  /** 打开透视表字段选择对话框,再从当前选区创建静态透视汇总表。需 `pivotTable` + `editable`。 */
  openPivotTableDialog(): boolean
  // ===== 条件格式编辑(1.9.0)需 `conditionalFormat` + `editable` =====
  /** 当前表的条件格式规则集(只读副本)。 */
  getConditionalRules(): ConditionalRule[]
  /** 新增一条规则(未给 id 自动派、origin 默认 'user');返回新 id 或 false。 */
  addConditionalRule(rule: Partial<ConditionalRule> & Pick<ConditionalRule, 'ranges' | 'type'>): string | false
  /** 按 id 改一条规则(浅合并)。 */
  updateConditionalRule(ruleId: string, patch: Partial<ConditionalRule>): boolean
  /** 按 id 删一条规则。 */
  removeConditionalRule(ruleId: string): boolean
  /** 整表替换规则集。 */
  setConditionalRules(rules: ConditionalRule[]): boolean
  /** 打开条件格式管理对话框(框架无关 DOM,三壳共用)。 */
  openConditionalFormatDialog(): boolean
  /** 给当前选区设数字格式代码(numFmt)。需 editable。1.11.0 */
  setSelectionNumberFormat(code: string): boolean
  /** 打开数字格式编辑对话框(框架无关 DOM,三壳共用)。需 editable + 选区。1.11.0 */
  openNumberFormatDialog(): boolean
  /** 启动格式刷:采样活动格样式,下次选区完成即刷上。sticky=连续刷。需 editable。1.12.0 */
  startFormatPainter(sticky?: boolean): boolean
  /** 格式刷是否待刷(工具栏 active 态)。1.12.0 */
  isFormatPainterArmed(): boolean
  /** 退出格式刷。1.12.0 */
  cancelFormatPainter(): void
  /** 读某格批注(无则 '')。1.11.0 */
  getCellComment(row: number, col: number): string
  /** 设/清某格批注(空串 = 删除)。需 editable。1.11.0 */
  setCellComment(row: number, col: number, comment: string): boolean
  /** 打开批注编辑对话框(默认活动格)。需 editable。1.11.0 */
  openCommentEditor(row?: number, col?: number): boolean
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
  /** 单元格"全息体检":snapshot + 合并区 + 浮动图覆盖 + WPS 内嵌图 + 数据验证 + 条件格式命中 + 链接/批注 */
  inspectCell(row: number, col: number): CellInspection | null
  /** 进入编辑(需有 editor 工厂 + 可编辑);返回是否进入 */
  beginEdit(row: number, col: number): boolean
  /** 取消当前编辑(不改模型) */
  cancelEdit(): void
  /** 当前是否有活动编辑器 */
  isEditing(): boolean
  /** 给区域套样式覆盖(E5;粗体/对齐/填充等);editable 时入命令栈(可撤销 + 发 cell-change + 记脏) */
  setStyle(range: MergeRange, patch: CellStyleOverride): boolean
  /** 活动格当前背景填充色(#RRGGBB;无填充→白) —— 工具栏色板回显用 */
  getActiveFillColor(): string
  /** 活动格当前字体色(#RRGGBB;缺省黑) */
  getActiveFontColor(): string
  /** 给当前选区设背景填充色(null=清除填充);editable 时入命令栈 */
  setSelectionFill(color: string | null): boolean
  /** 给当前选区设字体色;editable 时入命令栈 */
  setSelectionFontColor(color: string): boolean
  /** 当前选区里 wrapText 的整体态:'all' 全开 / 'none' 全关 / 'mixed' 混合(工具栏 active/右键勾选用) */
  getSelectionWrapState(): 'all' | 'none' | 'mixed'
  /** 切换当前选区的"自动换行"(WPS 风格 toggle);editable 时入命令栈,行高按内容重撑(只扩不缩) */
  toggleWrapTextOnSelection(): boolean
  /** 合并区域(G1;清空被覆盖格,只留左上锚点);editable 时入命令栈 */
  mergeCells(range: MergeRange): boolean
  /** 拆分区域内的合并(G1);editable 时入命令栈 */
  unmergeCells(range: MergeRange): boolean
  /** 把 TSV 文本粘到选区左上角(G2;类型自动推断、跳过只读、入命令栈);at 缺省用活动格 */
  pasteText(text: string, at?: { row: number; col: number }): boolean
  /** 解析 Excel/WPS 复制的剪贴板 HTML → 富粘贴(值+字体/颜色/填充/边框/对齐+合并+data-uri图),整体单次撤销。
   *  behaviorOverride = 逐次粘贴行为预设(右键「选择性粘贴」用;缺省走 setPasteBehavior 设的默认) */
  pasteRichHtml(html: string, at?: { row: number; col: number }, behaviorOverride?: Partial<PasteBehavior> | null): boolean
  /** 读当前粘贴行为配置(完整) */
  getPasteBehavior(): PasteBehavior
  /** 设粘贴行为默认(缺项回落默认);影响 Ctrl+V / 右键「粘贴」 */
  setPasteBehavior(cfg: Partial<PasteBehavior> | null): void
  /** 打开「粘贴行为配置」面板(框架无关 DOM,三壳共用);需 editable。返回是否打开 */
  openPasteConfigDialog(): boolean
  /** 把一张图片 blob 落到活动格(转内嵌图);剪贴板单图 / 拖文件进网格用 */
  pasteImageBlob(blob: Blob, at?: { row: number; col: number }): Promise<boolean>
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
  /** 活动格在公式栏里的可编辑字符串(公式→`=...`;数值→原始数字串;布尔→TRUE/FALSE;其余→显示文本) */
  getCellEditString(): string
  /** 活动格此刻是否可经公式栏/命令式编辑(editable 开 + 该格非只读) */
  canEditActiveCell(): boolean
  /** 经公式栏提交活动格的值(同 editCell 输入语义);move='down' 提交后活动格下移。返回是否提交 */
  commitActiveCellValue(value: string, move?: 'down'): boolean
  /** 读 WPS 单元格内嵌图(DISPIMG)登记表(id→{id,src,mime});非 WPS 文件返空数组 */
  getCellImages(): { id: string; src: string; mime?: string }[]
  /** 某格是否内嵌图 → {id,src,mime} 否则 null(供图片放大判定) */
  getCellImageAt(row: number, col: number): { id: string; src: string; mime?: string } | null
  /** 打开图片放大灯箱(命令式;src = blob/data/http url) */
  openImageLightbox(src: string, fileName?: string, mime?: string): void
  /** 设 WPS 单元格内嵌图贴合方式(fill 拉伸铺满 / contain 等比留白 / cover 等比裁剪);即时重绘 */
  setCellImageFit(fit: 'fill' | 'contain' | 'cover'): void
  /** 浮动图 → 单元格内嵌图(显式目标格);editable 时入命令栈 */
  convertImageToCell(imageIndex: number, row: number, col: number): boolean
  /** 浮动图 → 内嵌图(**就近**:图在哪格就嵌哪格,几何反推目标);editable 时入命令栈 */
  convertImageToCellAuto(imageIndex: number): boolean
  /** 批量把浮动图就近嵌入各自单元格(整表;`col` 给定则仅该列);一次进撤销栈,返回嵌入张数 */
  convertAllImagesToCells(col?: number): number
  /** 选区批量:把中心落在 range 内的浮动图全部就近嵌入,单次撤销;返回嵌入张数。
   *  壳侧 1.2.0 起返 `Promise<number>`(为接内置 ExportProgressOverlay;关闭 `:exportProgress` 也仍是 Promise) */
  convertImagesInRangeToCell(range: MergeRange): Promise<number>
  /** 选区批量(反向):range 内所有 DISPIMG 格拎成浮动图,单次撤销;返回转换张数(壳侧返 Promise,见上) */
  convertCellImagesInRangeToFloat(range: MergeRange, size?: { width: number; height: number }): Promise<number>
  /** 程序化打开右键菜单(Plan C;键盘 Shift+F10 / 工具栏触发等);items 不给则按当前选区算内置 */
  openContextMenu(x: number, y: number, items?: MenuItem[]): void
  /** 关闭当前打开的右键菜单 */
  closeContextMenu(): void
  /** 单元格内嵌图 → 浮动图(把 row,col 的 DISPIMG 拎成浮动图,默认 96×96px);editable 时入命令栈 */
  convertCellImageToFloat(row: number, col: number, size?: { width: number; height: number }): boolean
  /** 在 at 处插入 count 行(E7);editable 时入命令栈 + 发 struct-change */
  insertRows(at: number, count?: number): boolean
  /** 删除 [at, at+count) 行(与合并相交则相交合并被移除) */
  deleteRows(at: number, count?: number): boolean
  /** 在 at 处插入 count 列 */
  insertCols(at: number, count?: number): boolean
  /** 删除 [at, at+count) 列 */
  deleteCols(at: number, count?: number): boolean
  /**
   * 程序化设列宽 (px, 模型单位). Phase B 2026-06-08:
   * target 接 `number | number[] | {from,to}` (DimTarget). 多 index 时聚合成单次 undo.
   * 返回**成功条数** (0 = 全 skip / editable=false). 老 `setColumnWidth(5, 100)` 单值用法兼容.
   */
  setColumnWidth(target: import('./edit/types').DimTarget, width: number): number
  /** 程序化设行高. 同 setColumnWidth, 维度 = 'row'. */
  setRowHeight(target: import('./edit/types').DimTarget, height: number): number
  /** 批量 autoFit 列宽 (Phase B). target 不传 = 整表; 传 DimTarget = 选定列. 返成功条数. */
  autoFitColumns(target?: import('./edit/types').DimTarget): number
  /** 批量 autoFit 行高 (Phase B). 同上, 维度 = 'row'. */
  autoFitRows(target?: import('./edit/types').DimTarget): number
  /** 重置列宽到默认 (Phase B) — 移除 columns Map 条目, 回落 defaultColWidth. 返成功条数. */
  resetColumnWidth(target: import('./edit/types').DimTarget): number
  /** 重置行高到默认. 同上, 维度 = 'row'. */
  resetRowHeight(target: import('./edit/types').DimTarget): number
  /** 公式引擎是否已就绪(recalc 开启 + 异步 warm 完成);未开重算恒 false */
  isRecalcReady(): boolean
  /** 当前虚拟范围(滚动出空行/列的外推上限,含 dimension 兜底);不动 dimension/文件 */
  getVirtualExtent(): { rows: number; cols: number }
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
  /** 右键菜单 transform(Plan C):`(ctx, items) => items[]`,多插件按数组顺序串行,组件 :contextMenu prop 最后覆盖 */
  contextMenu?: import('./viewer/controller').ContextMenuTransform
  /** 高级: 拿命令式 API、订阅事件;返回可选清理函数 */
  setup?: (ctx: ExcelPluginContext) => void | (() => void)
}

/** 定义插件(仅作类型推断,原样返回) */
export function definePlugin(plugin: ExcelPlugin): ExcelPlugin {
  return plugin
}

// ---- P3 公开导出:JSON 直渲 + 模板样式 overlay(给"仅引擎"用户) ----
export { jsonToWorkbook, isWorkbookModel, type JsonInput, type JsonLoadOptions, type JsonRow, type JsonSheetInput } from './loader-json'
export { applyStyleTemplate } from './template/style-overlay'
export type { CellInspection } from './model/inspect'
export type { EditableTarget, EditConfig } from './edit/types'
export type { MenuItem } from './edit/context-menu'
export type {
  ContextMenuCtx,
  ContextMenuTransform,
  ContextMenuBeforePayload,
  ContextMenuShowPayload,
} from './viewer/controller'
