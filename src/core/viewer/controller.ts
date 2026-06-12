/**
 * ViewerController(框架无关)—— 接管"交互式网格"的渲染引擎 + 选区 + 交互,供 Vue / React 壳共用。
 *
 * 职责: renderer 生命周期、view 状态(滚动/缩放/尺寸)、render 调度(rAF)、measure、spacer 尺寸、
 * 列宽行高拖拽/自适应、叠加层(OverlayManager)、几何 API(rectOf);
 * 选区模型(anchor/active/mode)、鼠标/键盘交互、命中检测、悬停 tooltip、复制、滚动到视图。
 * find/filter 仍在壳里(通过 renderer + 本控制器选区方法操作),A2c 再下沉。
 *
 * 与框架的桥接全走 hooks: onRenderer(把 renderer 镜像回壳的响应式)、onRenderTick(壳据此重算 overlay slot)、
 * onSelectionChange(壳据此 +1 让选区相关计算属性重算)、onCellClick/onCellDblClick/onHyperlink/onFilterButton/onTooltip
 * (交互回调,壳决定 emit / 插件派发 / 策略)。壳不需镜像 contentSize/selection —— 直接读控制器。
 */
import type { CellModel, CellStyle, CellStyleOverride, ColumnInfo, ImageAnchor, MergeRange, PivotFilterRule, PivotSummary, PivotTableLayout, PivotTableModel, RowInfo, SheetModel, WorkbookModel } from '../model/types'
import { cellKey } from '../model/types'
import { setCellValue, setImageRect, cloneImageAnchor } from '../model/mutations'
import { pxToEmu, emuToPx } from '../layout/units'
import { deleteIntersectsMerge } from '../model/structure'
import { anchorRect } from '../overlay/anchor'
import type { EditConfig } from '../edit/types'
import { resolveEditable, canEditDimension, normalizeDimTarget } from '../edit/permissions'
import { cloneWorkbook } from '../model/clone'
import { EditController, type EditControllerHost, type EditEventName } from '../edit/edit-controller'
import { defaultFormulaEngineFactory } from '../formula/hyperformula-adapter'
import type { CellValue, SheetToJSONOptions } from '../model/data-access'
import type { CellSnapshot } from '../model/snapshot'
import { inspectCell, type CellInspection } from '../model/inspect'
// 模板样式 overlay 在壳层做(controller 不直接持有),保留 import 给类型链或后续重新接入用
// import { applyStyleTemplate } from '../template/style-overlay'
import { CellEditorHost } from '../edit/editor-host'
import { ContextMenuHost, type MenuItem } from '../edit/context-menu'
import { PivotDialogHost, PivotFieldPanelHost, type PivotFieldOption, type PivotOutputChoice } from './pivot-dialog-host'
import { PasteConfigDialogHost } from './paste-config-host'
import { ReadOnlyPromptHost } from './readonly-prompt-host'
import { parseClipboardHtml } from '../edit/clipboard-html'
import { serializeSnapshot, encodeSnapshot, parseSnapshotHtml, withoutImages, bytesToB64, CLIP_IMAGE_BUDGET_BYTES } from '../edit/clipboard-snapshot'
import { type PasteBehavior, DEFAULT_PASTE_BEHAVIOR, PASTE_PRESET_VALUES_ONLY } from '../edit/paste-behavior'
import { LightboxHost } from './lightbox-host'
import type { CellEditorContext, EditorCommitValue, EditorResolver } from '../edit/editor-context'
import { defaultCellEditor } from '../edit/default-editor'
import { CanvasRenderer, type CellImageFit, type RendererOptions, type ViewState } from '../render/canvas-renderer'
import { MAX_GRID_ROWS, MAX_GRID_COLS } from '../layout/grid-metrics'
import { invalidateAutofit } from '../layout/autofit'
import { toHex6 } from '../format/color'
import { OverlayManager, type OverlayQuads } from './overlay-manager'
import { WorkbookExporter, type ExporterHost } from '../export/exporter'
import type { ImageExportOptions, PdfExportOptions, PrintOptions } from '../export/types'
import type { XlsxExportOptions } from '../export/xlsx-writer'
import type { CreatePivotTableOptions, PermissionDeniedPayload } from '../plugin'

export type Cell = { row: number; col: number }
export interface TooltipState {
  text: string
  x: number
  y: number
  kind: 'overflow' | 'comment'
}

/** 查找状态快照(供壳渲染 FindBar) */
export interface FindState {
  query: string
  matchCase: boolean
  wholeCell: boolean
  count: number
  index: number
}

/** 自动筛选下拉浮层(列去重值 + 已选 + 屏幕位置) */
export interface FilterPopupState {
  col: number
  values: string[]
  selected: string[]
  x: number
  y: number
  /** 该列当前排序方向(用于浮层高亮 ↑/↓);null = 未按此列排序 */
  sortDir: 'asc' | 'desc' | null
}

export interface ViewerControllerEls {
  canvas: HTMLCanvasElement
  /** 量可视区尺寸用 */
  renderArea: HTMLElement
  /** 原生滚动容器 */
  scroller: HTMLElement
  /** 撑开滚动范围的占位元素(控制器直接设其宽高) */
  spacer: HTMLElement
  /** 叠加层四象限 */
  overlays: OverlayQuads
  /** 单元格编辑器挂载层(在格 + overlay 之上;E2) */
  editorSlot: HTMLElement
}

export interface ViewerControllerHooks {
  /** renderer 重建时回调,壳据此镜像到响应式(保持现有 renderer.value 读法 + chrome 响应) */
  onRenderer: (renderer: CanvasRenderer | null) => void
  /** 每次绘制后回调,壳据此 +1 让 overlay slot 重算位置 */
  onRenderTick: () => void
  /** 选区模型变化(壳据此 +1 让 selection/activeCell/stats 等计算属性重算) */
  onSelectionChange: () => void
  /** 单击单元格(壳 emit 'cell-click' + 派发插件) */
  onCellClick: (row: number, col: number, text: string) => void
  /** 双击单元格(壳 emit 'cell-dblclick') */
  onCellDblClick: (row: number, col: number, text: string) => void
  /** 单击超链接(壳 emit 'hyperlink-click' + 按 openLinks 决定是否 window.open) */
  onHyperlink: (url: string, cell: Cell) => void
  /** 悬停提示变化(壳渲染 tooltip DOM) */
  onTooltip: (tip: TooltipState | null) => void
  /** 查找状态变化(壳据此 +1 让 FindBar / 工具栏重算) */
  onFindChange: () => void
  /** 筛选状态/浮层变化(壳据此 +1 让 FilterPopup / 工具栏重算) */
  onFilterChange: () => void
  /** core 里新增/切换工作表时通知壳同步 activeSheet。 */
  onActiveSheetChange?: (index: number) => void
  /** 编辑事件(cell-change/edit-start/edit-commit;壳转 emit + 插件派发) */
  onEditEvent: (event: EditEventName, payload: unknown) => void
  /** 右键菜单触发前(Plan C):用户可调 `preventDefault()` 阻止内置菜单弹出(然后自渲染) */
  onContextMenuBefore?: (payload: ContextMenuBeforePayload) => void
  /** 右键菜单"展示"通知(Plan C):无论内置是否弹都触发,供壳自渲染或事件流串到业务 */
  onContextMenuShow?: (payload: ContextMenuShowPayload) => void
}

/** 右键菜单上下文(单格/选区 + 活动格 + 当前表/簿 + editable 闸门态) */
export interface ContextMenuCtx {
  range: MergeRange
  single: boolean
  activeCell: Cell
  sheet: SheetModel
  workbook: WorkbookModel
  editable: boolean
}
/** 用户 transform:在内置 items 上做加 / 减 / 重排;返新数组生效,返 undefined 用原样 */
export type ContextMenuTransform = (ctx: ContextMenuCtx, items: MenuItem[]) => MenuItem[] | undefined | void
export interface ContextMenuBeforePayload {
  event: MouseEvent
  ctx: ContextMenuCtx
  items: MenuItem[]
  preventDefault: () => void
}
export interface ContextMenuShowPayload {
  x: number
  y: number
  ctx: ContextMenuCtx
  items: MenuItem[]
}

const BLANK = '(空白)'

export class ViewerController {
  /** 当前渲染器(壳通过 onRenderer 镜像) */
  renderer: CanvasRenderer | null = null
  /** 视图状态(滚动/缩放/尺寸);壳与控制器共享同一对象 */
  readonly view: ViewState = { scrollX: 0, scrollY: 0, width: 0, height: 0, zoom: 1 }

  private overlays: OverlayManager
  private sheet: SheetModel | null = null
  private rafId = 0
  private contentW = 0
  private contentH = 0
  /** 虚拟外推行/列数(滚动出空行/空列;只增不减,封顶 Excel 上限)。不动 dimension/文件。 */
  private virtualRows = 0
  private virtualCols = 0

  // ---- 导出上下文(供 WorkbookExporter 取数) ----
  private workbook: WorkbookModel | null = null
  private activeIndex = 0
  private rendererOpts: RendererOptions = {}
  /** 下载默认文件名(壳可随 props 更新) */
  fileName: string | undefined = undefined
  /** 右键上下文菜单宿主(G3;body 级 DOM,框架无关) */
  private menuHost = new ContextMenuHost()
  private lightbox = new LightboxHost()
  private pivotDialog = new PivotDialogHost()
  private pivotPanel = new PivotFieldPanelHost()
  private pasteConfigDialog = new PasteConfigDialogHost()
  private readonlyPrompt = new ReadOnlyPromptHost()
  /** 透视表"活刷新"重入兜底:recompute 内部 setCellValue 不再触发 onModelChange,此旗保险。 */
  private pivotRefreshing = false
  private lightboxEnabled = true
  /** 用户的右键菜单 transform 回调(Plan C):`(ctx, items) => MenuItem[] | undefined` */
  private ctxMenuTransform: ContextMenuTransform | null = null
  /** 原始 .xlsx 字节(壳加载时注入;供高保真 overlay 导出重载原件) */
  private sourceBuffer: ArrayBuffer | null = null
  /** 壳在加载后注入原始字节(供 exportXlsx({fidelity:'overlay'}) 重载原件叠加编辑) */
  setSourceBuffer(buf: ArrayBuffer | null): void {
    this.sourceBuffer = buf
  }
  private exporter: WorkbookExporter

  // ---- 选区模型 ----
  private selAnchor: Cell | null = null // 固定角(扩选时不动)
  private selActive: Cell | null = null // 活动角(移动/扩选时变)
  private selMode: 'range' | 'rows' | 'cols' = 'range'

  // ---- 拖拽态 ----
  private dragMode: 'none' | 'cell' | 'row' | 'col' | 'resize-col' | 'resize-row' | 'image' = 'none'
  private resizeTarget = -1 // 正在拖拽改宽高的列/行索引
  // 图片拖拽态(E6)
  private imageDragIdx = -1
  private imageDragStartRect: { left: number; top: number; width: number; height: number } | null = null
  private imageDragStartMouse = { x: 0, y: 0 }
  private imageDragBefore: ImageAnchor | null = null
  private resizeStartPos = 0 // 起始鼠标坐标(px)
  private resizeStartSize = 0 // 起始宽/高(px,zoom 后)
  private resizeStartInfo: ColumnInfo | RowInfo | null = null // 拖拽起始的列/行维度信息(克隆,供 undo)
  private resizeStartModelSize = 0 // 拖拽起始的模型 px 尺寸(非缩放,dim-change 事件用)
  private dragMoved = false
  private dragStartXY: { x: number; y: number } | null = null // mousedown 起点(死区判定:超 3px 才算拖动)

  // ---- 查找态 ----
  private findQuery = ''
  private findMatchCase = false
  private findWholeCell = false
  private findHits: Cell[] = []
  private findIndex = -1

  // ---- 自动筛选态(仅作用于当前表) ----
  private filterState = new Map<number, Set<string>>() // 列 → 允许值集合(缺省=未筛选)
  private filterOrigHidden = new Map<number, boolean>() // 行 → 原始 hidden(首次筛选前快照)
  private filterPopup: FilterPopupState | null = null

  // ---- 排序态(仅作用于当前表;rebuild 时重置) ----
  private sortCol = -1
  private sortDir: 'asc' | 'desc' | null = null

  // ---- 编辑配置(默认只读;E0 只做闸门,后续阶段在此扩展) ----
  private editCfg: EditConfig = {}
  /** 编辑底座(命令栈/快照/事件;E1) */
  readonly edit: EditController
  /** 单元格编辑器宿主(E2) */
  private editorHost: CellEditorHost
  /** 按格解析编辑器(壳合并 plugin.editor + prop.editor 后注入;E2) */
  private editorResolver?: EditorResolver
  /** paste 事件处理器(绑在 scroller 上;Ctrl+V 走它拿**原始**剪贴板 HTML,避开 clipboard.read() 的净化) */
  private readonly onPasteHandler = (e: ClipboardEvent) => this.onPaste(e)

  constructor(
    private els: ViewerControllerEls,
    private hooks: ViewerControllerHooks,
  ) {
    this.overlays = new OverlayManager(els.overlays)
    const host: ExporterHost = {
      getWorkbook: () => this.workbook,
      getActiveIndex: () => this.activeIndex,
      getLiveRenderer: () => this.renderer,
      getRendererOpts: () => this.rendererOpts,
      getFileName: () => this.fileName,
      getSourceBuffer: () => this.sourceBuffer,
      isPivotEnabled: () => !!this.editCfg.pivotTable,
    }
    this.exporter = new WorkbookExporter(host)

    const editHost: EditControllerHost = {
      getSheet: () => this.sheet,
      getWorkbook: () => this.workbook,
      getDate1904: () => this.workbook?.date1904 ?? false,
      isEditable: (row, col) => this.isCellEditable(row, col),
      isEditingEnabled: () => !!this.editCfg.editable,
      getActiveSheetIndex: () => this.activeIndex,
      isRecalcEnabled: () => !!this.editCfg.editable && !!this.editCfg.recalc,
      getEngineFactory: () =>
        this.editCfg.editable && this.editCfg.recalc
          ? (this.editCfg.formulaEngine ?? defaultFormulaEngineFactory)
          : null,
      onModelChange: () => {
        this.refreshPivotsAfterEdit() // 源数据改动 → 透视表"活刷新"(在重算 metrics 之前改完模型)
        this.renderer?.rebuildMetrics()
        this.refreshContentSize()
        this.render()
      },
      rebuildOverlays: () => {
        if (this.sheet && this.renderer) void this.overlays.build(this.sheet, this.renderer, this.view)
      },
      emit: (event, payload) => this.emitEditEvent(event, payload),
    }
    this.edit = new EditController(editHost)
    this.editorHost = new CellEditorHost(els.editorSlot, (row, col) => this.rectOf(row, col))
    // Ctrl+V 走 paste 事件:e.clipboardData 给的是原始未净化 HTML(WPS 的 <style>/VML 都在);
    // clipboard.read()(pasteFromClipboard 用)会净化删掉 <style>/注释 → 丢格式/数字格式/内嵌图。
    els.scroller.addEventListener('paste', this.onPasteHandler)
  }

  /** 切表/换簿/主题变化: 清状态,重建渲染器,重置滚动,量尺寸,建叠加层,绘制,按需重跑查找 */
  rebuild(sheet: SheetModel, workbook: WorkbookModel, zoom: number, opts: RendererOptions): void {
    // 先清状态(选区作废;查找命中作废但保留 query 以便新表重跑;tooltip 隐藏)
    this.selAnchor = null
    this.selActive = null
    this.selMode = 'range'
    this.hooks.onSelectionChange()
    this.findHits = []
    this.findIndex = -1
    this.hooks.onTooltip(null)
    this.sortCol = -1
    this.sortDir = null
    this.editorHost.unmount() // 卸掉活动编辑器
    this.edit.reset() // 切表/换簿:命令栈 + 编辑态作废
    const freshWorkbook = this.workbook !== workbook // 换新簿 vs 仅切表

    this.sheet = sheet
    this.workbook = workbook
    if (freshWorkbook) {
      this.edit.resetDirtyBaseline() // 换新簿:作废 baseline + 清脏(切表保留)
      this.edit.refreshEngine() // 换新簿:为新簿重建公式引擎(切表不重建)
    }
    this.activeIndex = Math.max(0, workbook.sheets.indexOf(sheet))
    this.rendererOpts = opts
    this.renderer = new CanvasRenderer(this.els.canvas, sheet, workbook, zoom, {
      ...opts,
      onNeedsRedraw: () => this.scheduleRender(), // WPS 内嵌图异步解码完触发重绘
      isEditable: (r, c) => this.isCellEditable(r, c), // Phase C 2026-06-08: 让渲染层感知 editable
    })
    this.hooks.onRenderer(this.renderer)
    this.view.zoom = zoom
    this.view.scrollX = 0
    this.view.scrollY = 0
    this.virtualRows = 0 // 换表/换簿:虚拟范围归零,measure() 据新视口重算
    this.virtualCols = 0
    this.els.scroller.scrollLeft = 0
    this.els.scroller.scrollTop = 0
    this.refreshContentSize()
    this.measure()
    this.overlays.build(sheet, this.renderer, this.view)
    this.render()
    if (this.findQuery) this.recomputeFind() // 新表上重跑当前查找
  }

  /** 立即绘制(取消挂起的 rAF)。绘制 + 定位叠加层 + 通知 tick。 */
  render(): void {
    if (this.rafId) {
      cancelAnimationFrame(this.rafId)
      this.rafId = 0
    }
    const r = this.renderer
    if (!r) return
    r.setSelection(this.getSelection())
    r.render(this.view)
    if (this.sheet) this.overlays.position(this.sheet, r, this.view)
    this.editorHost.position() // 活动编辑器随滚动/缩放跟随(无则 no-op)
    this.hooks.onRenderTick()
  }

  /** 合并到下一帧绘制(滚动/拖选高频时,每帧最多一次) */
  scheduleRender(): void {
    if (this.rafId) return
    this.rafId = requestAnimationFrame(() => {
      this.rafId = 0
      this.render()
    })
  }

  /** 量可视区尺寸 */
  measure(): void {
    this.view.width = this.els.renderArea.clientWidth
    this.view.height = this.els.renderArea.clientHeight
    this.recomputeVirtualExtent() // 视口变 → 至少留一屏空行/列可滚
  }

  /** 同步滚动量(壳的 scroll 事件里调) */
  setScroll(scrollX: number, scrollY: number): void {
    this.view.scrollX = scrollX
    this.view.scrollY = scrollY
    this.recomputeVirtualExtent() // 滚到底自动延伸出更多空行/列(spacer 增长 → 无限感)
    this.scheduleRender()
  }

  /** 缩放: 保持视口中心相对内容的比例,避免跳到左上角 */
  setZoom(zoom: number): void {
    const r = this.renderer
    const sc = this.els.scroller
    if (!r) return
    const ratioX = this.contentW ? (sc.scrollLeft + sc.clientWidth / 2) / this.contentW : 0
    const ratioY = this.contentH ? (sc.scrollTop + sc.clientHeight / 2) / this.contentH : 0
    r.setZoom(zoom)
    this.view.zoom = zoom
    this.recomputeVirtualExtent() // 缩放改变可见行列数 → 虚拟范围按需延伸
    this.refreshContentSize() // 先把 spacer 撑到新尺寸,再设滚动(无需等框架 tick)
    sc.scrollLeft = Math.max(0, ratioX * this.contentW - sc.clientWidth / 2)
    sc.scrollTop = Math.max(0, ratioY * this.contentH - sc.clientHeight / 2)
    this.view.scrollX = sc.scrollLeft
    this.view.scrollY = sc.scrollTop
    this.render()
  }

  /** 拖拽改列宽(px) */
  setColWidthPx(col: number, px: number): void {
    this.renderer?.setColWidthPx(col, px)
    this.refreshContentSize()
    this.scheduleRender()
  }
  /** 拖拽改行高(px) */
  setRowHeightPx(row: number, px: number): void {
    this.renderer?.setRowHeightPx(row, px)
    this.refreshContentSize()
    this.scheduleRender()
  }
  /** 双击列边界: 自适应列宽(editable 时入命令栈/发事件/记脏) */
  autoFitColumn(col: number): void {
    if (this.editCfg.editable) this.beginResizeRecord('col', col)
    this.renderer?.autoFitColumn(col)
    this.refreshContentSize()
    this.render()
    if (this.editCfg.editable) this.endResizeRecord('col', col)
  }
  /** 双击行边界: 自适应行高(editable 时入命令栈/发事件/记脏) */
  autoFitRow(row: number): void {
    if (this.editCfg.editable) this.beginResizeRecord('row', row)
    this.renderer?.autoFitRow(row)
    this.refreshContentSize()
    this.render()
    if (this.editCfg.editable) this.endResizeRecord('row', row)
  }

  // ---- 维度 / 脏状态 命令式 API(E3.5;Phase B 多形态 2026-06-08) ----
  /**
   * 程序化设列宽 (px, 模型单位/非缩放). Phase B 2026-06-08:
   * target 接 `number | number[] | {from,to}` (DimTarget). 多 index 时聚合成单次 undo.
   * editable 时走命令栈;strictDimensions=true 时该列至少 1 格在白名单内才生效.
   * 返回**成功条数** (0 = 全部 skip / editable=false).
   */
  setColumnWidth(target: import('../edit/types').DimTarget, width: number): number {
    const sheet = this.sheet
    if (!sheet) return 0
    const indices = normalizeDimTarget(target)
    return this.edit.setDimensions('col', indices, width, (i) => canEditDimension(sheet, 'col', i, this.editCfg))
  }
  /** 程序化设行高 (px, 模型单位/非缩放). 同 setColumnWidth, 维度 = 'row'. */
  setRowHeight(target: import('../edit/types').DimTarget, height: number): number {
    const sheet = this.sheet
    if (!sheet) return 0
    const indices = normalizeDimTarget(target)
    return this.edit.setDimensions('row', indices, height, (i) => canEditDimension(sheet, 'row', i, this.editCfg))
  }
  /**
   * 批量 autoFit 列宽 (Phase B 2026-06-08). target 不传 = 整表; 传 DimTarget = 选定列.
   * 单 index 走 autoFitColumn (含 resize-record); 多 index 单次 restore-wb undo + 循环 autofit.
   * 返回成功条数.
   */
  autoFitColumns(target?: import('../edit/types').DimTarget): number {
    const sheet = this.sheet
    const r = this.renderer
    if (!sheet || !r) return 0
    const indices = target === undefined ? Array.from({ length: sheet.dimension.cols }, (_, i) => i) : normalizeDimTarget(target)
    const allowed = indices.filter((i) => i >= 0 && canEditDimension(sheet, 'col', i, this.editCfg))
    const denied = indices.filter((i) => i >= 0 && !canEditDimension(sheet, 'col', i, this.editCfg))
    if (denied.length) {
      this.hooks.onEditEvent('permission-denied', { reason: 'dimension', cells: [], dims: { axis: 'col', indices: denied }, message: `${denied.length} 列未覆盖白名单,autoFit 跳过` })
    }
    if (!allowed.length) return 0
    if (allowed.length === 1) {
      this.autoFitColumn(allowed[0])
      return 1
    }
    // 多列: 单次快照, 循环 autoFit (renderer.autoFitColumn 内部直接写 sheet.columns); 一次 undo
    const wb = this.workbook
    if (!wb) return 0
    this.edit.ensureBaseline()
    const snap = cloneWorkbook(wb)
    let written = 0
    for (const i of allowed) {
      r.autoFitColumn(i)
      written++
    }
    this.edit.pushUndoExternal({ kind: 'restore-wb', snapshot: snap })
    this.edit.markDirtyExternal()
    this.refreshContentSize()
    this.render()
    return written
  }
  /** 批量 autoFit 行高 (Phase B 2026-06-08). 同 autoFitColumns, 维度 = 'row'. */
  autoFitRows(target?: import('../edit/types').DimTarget): number {
    const sheet = this.sheet
    const r = this.renderer
    if (!sheet || !r) return 0
    const indices = target === undefined ? Array.from({ length: sheet.dimension.rows }, (_, i) => i) : normalizeDimTarget(target)
    const allowed = indices.filter((i) => i >= 0 && canEditDimension(sheet, 'row', i, this.editCfg))
    const denied = indices.filter((i) => i >= 0 && !canEditDimension(sheet, 'row', i, this.editCfg))
    if (denied.length) {
      this.hooks.onEditEvent('permission-denied', { reason: 'dimension', cells: [], dims: { axis: 'row', indices: denied }, message: `${denied.length} 行未覆盖白名单,autoFit 跳过` })
    }
    if (!allowed.length) return 0
    if (allowed.length === 1) {
      this.autoFitRow(allowed[0])
      return 1
    }
    const wb = this.workbook
    if (!wb) return 0
    this.edit.ensureBaseline()
    const snap = cloneWorkbook(wb)
    let written = 0
    for (const i of allowed) {
      r.autoFitRow(i)
      written++
    }
    this.edit.pushUndoExternal({ kind: 'restore-wb', snapshot: snap })
    this.edit.markDirtyExternal()
    this.refreshContentSize()
    this.render()
    return written
  }
  /**
   * 重置列宽到默认 (Phase B 2026-06-08) — 移除 sheet.columns Map 条目, 回落到 defaultColWidth.
   * 多 index 单次 undo. 返回成功条数.
   */
  resetColumnWidth(target: import('../edit/types').DimTarget): number {
    const sheet = this.sheet
    if (!sheet) return 0
    const indices = normalizeDimTarget(target)
    return this.edit.resetDimensions('col', indices, (i) => canEditDimension(sheet, 'col', i, this.editCfg))
  }
  /** 重置行高到默认. 同 resetColumnWidth, 维度 = 'row'. */
  resetRowHeight(target: import('../edit/types').DimTarget): number {
    const sheet = this.sheet
    if (!sheet) return 0
    const indices = normalizeDimTarget(target)
    return this.edit.resetDimensions('row', indices, (i) => canEditDimension(sheet, 'row', i, this.editCfg))
  }
  /** 公式引擎是否已就绪(recalc 开启 + 异步 warm 完成);未开重算恒 false。 */
  isRecalcReady(): boolean {
    return this.edit.isRecalcReady()
  }
  /** 当前是否有未保存修改(自加载/还原以来发生过编辑或 resize)。 */
  isDirty(): boolean {
    return this.edit.isDirty()
  }
  /** 放弃全部修改,还原到刚加载的原件;返回是否还原(无 baseline → false)。 */
  resetToOriginal(): boolean {
    return this.edit.resetToOriginal()
  }

  /** 单元格当前屏幕矩形(render-area 相对) */
  rectOf(row: number, col: number): { x: number; y: number; w: number; h: number } | null {
    const r = this.renderer
    return r ? r.screenRectOfCell(this.view, row, col) : null
  }
  /** 区域当前屏幕矩形(左上到右下并集) */
  rectOfRange(range: MergeRange): { x: number; y: number; w: number; h: number } | null {
    const r = this.renderer
    if (!r) return null
    const tl = r.screenRectOfCell(this.view, range.top, range.left)
    const br = r.screenRectOfCell(this.view, range.bottom, range.right)
    return { x: tl.x, y: tl.y, w: br.x + br.w - tl.x, h: br.y + br.h - tl.y }
  }

  dispose(): void {
    if (this.rafId) cancelAnimationFrame(this.rafId)
    this.rafId = 0
    this.els.scroller.removeEventListener('paste', this.onPasteHandler)
    this.overlays.dispose()
    this.editorHost.dispose()
    this.menuHost.dispose()
    this.lightbox.dispose()
    this.pivotDialog.dispose()
    this.pivotPanel.dispose()
    this.pasteConfigDialog.dispose()
    this.readonlyPrompt.dispose()
    this.edit.disposeEngine()
  }

  /** 内容总尺寸变化 → 量 + 直接撑 spacer(滚动范围由它决定) */
  refreshContentSize(): void {
    const r = this.renderer
    if (!r) return
    this.contentW = r.contentWidth
    this.contentH = r.contentHeight
    this.els.spacer.style.width = this.contentW + 'px'
    this.els.spacer.style.height = this.contentH + 'px'
  }

  /**
   * 据当前滚动+视口算"虚拟范围"(滚到数据下方仍有空行/空列可滚动/选中/编辑)。
   * 只增不减(防 spacer 抖动)、封顶 Excel 上限。**不动 dimension/文件**;编辑虚拟格才靠 growDimension 变实。
   * 仅当范围真变化时重建 GridMetrics + 刷 spacer。
   */
  recomputeVirtualExtent(): void {
    const r = this.renderer
    if (!r) return
    const m = r.metrics
    // 视口可见内容底部/右缘落在哪行哪列,再各加一屏缓冲(滚到底时已经有下一屏空行)
    const bottomRow = m.rowAt(this.view.scrollY + this.view.height - m.colHeaderHeight) + 30
    const rightCol = m.colAt(this.view.scrollX + this.view.width - m.rowHeaderWidth) + 10
    const nextRows = Math.min(MAX_GRID_ROWS, Math.max(this.virtualRows, bottomRow))
    const nextCols = Math.min(MAX_GRID_COLS, Math.max(this.virtualCols, rightCol))
    if (nextRows === this.virtualRows && nextCols === this.virtualCols) return
    this.virtualRows = nextRows
    this.virtualCols = nextCols
    if (r.setVirtualExtent(nextRows, nextCols)) this.refreshContentSize()
  }

  /** 当前虚拟范围(含 dimension 兜底);供调试/e2e。 */
  getVirtualExtent(): { rows: number; cols: number } {
    const m = this.renderer?.metrics
    return { rows: m?.vRows ?? 0, cols: m?.vCols ?? 0 }
  }

  // ====================== 选区模型 ======================

  /** 当前选区(随 mode 解析为单元格/整行/整列范围,含合并单元格扩展) */
  getSelection(): MergeRange | null {
    const r = this.renderer
    const a = this.selAnchor
    const b = this.selActive
    if (!r || !a || !b) return null
    if (this.selMode === 'rows') {
      return { top: Math.min(a.row, b.row), bottom: Math.max(a.row, b.row), left: 0, right: r.metrics.cols - 1 }
    }
    if (this.selMode === 'cols') {
      return { left: Math.min(a.col, b.col), right: Math.max(a.col, b.col), top: 0, bottom: r.metrics.rows - 1 }
    }
    const ra = r.mergeAt(a.row, a.col) ?? cellRange(a)
    const rb = r.mergeAt(b.row, b.col) ?? cellRange(b)
    return {
      top: Math.min(ra.top, rb.top),
      left: Math.min(ra.left, rb.left),
      bottom: Math.max(ra.bottom, rb.bottom),
      right: Math.max(ra.right, rb.right),
    }
  }

  /** 活动单元格(选区的"焦点"角) */
  getActiveCell(): Cell | null {
    return this.selActive
  }

  /**
   * 活动格(或指定格)在公式栏里**可编辑的字符串**:
   * 公式 → `=...`;数值 → 原始数字串(非格式化,避免编辑货币/千分位被当文本);布尔 → TRUE/FALSE;
   * 日期/字符串/富文本 → 显示文本。空格返 ''。供公式栏 / 命令式取活动格编辑值。
   */
  getCellEditString(row?: number, col?: number): string {
    const a = row != null && col != null ? { row, col } : this.selActive
    if (!a || !this.sheet) return ''
    const cell = this.sheet.cells.get(cellKey(a.row, a.col))
    if (!cell) return ''
    if (cell.formula) return '=' + cell.formula
    const raw = cell.raw
    if (typeof raw === 'number') return String(raw)
    if (typeof raw === 'boolean') return raw ? 'TRUE' : 'FALSE'
    return this.renderer?.cellText(a.row, a.col) ?? (raw == null ? '' : String(raw))
  }

  /** 活动格此刻是否可经公式栏/命令式编辑(editable 开 + 该格非只读) */
  canEditActiveCell(): boolean {
    const a = this.selActive
    return !!(a && this.editCfg.editable && this.isCellEditable(a.row, a.col))
  }

  /**
   * 经公式栏提交活动格的值(value 同 editCell 的输入语义:`=`→公式、数字串→数字…)。
   * 仅在 editable + 该格可编辑时生效;move='down' 时提交后活动格下移(像 Excel 回车)。返回是否提交。
   */
  commitActiveCellValue(value: string, move?: 'down'): boolean {
    const a = this.selActive
    if (!a || !this.editCfg.editable || !this.isCellEditable(a.row, a.col)) return false
    // 若正有内嵌编辑器开着,先取消(避免双重提交打架)
    if (this.editorHost.isActive()) this.cancelEdit()
    // 仅当值真变化才入命令栈(避免把格式化显示文本当字符串回写、避免空提交污染 undo 栈)
    if (value !== this.getCellEditString(a.row, a.col)) this.edit.editCell(a.row, a.col, value)
    if (move === 'down' && this.renderer) {
      const nr = Math.min(a.row + 1, this.renderer.metrics.rows - 1)
      this.selectCell(nr, a.col) // selectCell 内含滚动到视图
    }
    return true
  }

  /** 清空选区 */
  clearSelection(): void {
    this.selAnchor = null
    this.selActive = null
    this.selMode = 'range'
    this.hooks.onSelectionChange()
  }

  /** 全选 */
  selectAll(): void {
    const r = this.renderer
    if (!r) return
    this.selMode = 'range'
    this.selAnchor = { row: 0, col: 0 }
    this.selActive = { row: r.metrics.rows - 1, col: r.metrics.cols - 1 }
    this.hooks.onSelectionChange()
  }

  /** 选中单个单元格并滚动到视图(查找定位用);range 模式,anchor=active */
  selectCell(row: number, col: number): void {
    this.selMode = 'range'
    this.selAnchor = { row, col }
    this.selActive = { row, col }
    this.scrollActiveIntoView()
    this.hooks.onSelectionChange()
    this.render()
  }

  /** 滚动到指定单元格;select=true 时同步选中目标格。 */
  scrollToCell(row: number, col: number, opts?: { select?: boolean }): boolean {
    const r = this.renderer
    if (!r) return false
    const targetRow = Math.max(0, Math.min(row, MAX_GRID_ROWS - 1))
    const targetCol = Math.max(0, Math.min(col, MAX_GRID_COLS - 1))
    if (targetRow >= r.metrics.vRows || targetCol >= r.metrics.vCols) {
      this.virtualRows = Math.max(this.virtualRows, targetRow + 1)
      this.virtualCols = Math.max(this.virtualCols, targetCol + 1)
      if (r.setVirtualExtent(this.virtualRows, this.virtualCols)) this.refreshContentSize()
    }
    const m = r.mergeAt(targetRow, targetCol)
    const target = m ? { row: m.top, col: m.left } : { row: targetRow, col: targetCol }
    if (opts?.select) {
      this.selectCell(target.row, target.col)
      return true
    }
    this.scrollCellIntoView(target)
    this.render()
    return true
  }

  /** 命令式设选区(anchor=左上, active=右下) */
  setSelectionRange(range: MergeRange): void {
    this.selMode = 'range'
    this.selAnchor = { row: range.top, col: range.left }
    this.selActive = { row: range.bottom, col: range.right }
    this.hooks.onSelectionChange()
    this.render()
  }

  private setCell(cell: Cell, extend: boolean): void {
    this.selMode = 'range'
    if (extend && this.selAnchor) this.selActive = cell
    else {
      this.selAnchor = cell
      this.selActive = cell
    }
    this.hooks.onSelectionChange()
  }
  private setRows(row: number, extend: boolean): void {
    this.selMode = 'rows'
    const c: Cell = { row, col: 0 }
    if (extend && this.selAnchor) this.selActive = c
    else {
      this.selAnchor = c
      this.selActive = c
    }
    this.hooks.onSelectionChange()
  }
  private setCols(col: number, extend: boolean): void {
    this.selMode = 'cols'
    const c: Cell = { row: 0, col }
    if (extend && this.selAnchor) this.selActive = c
    else {
      this.selAnchor = c
      this.selActive = c
    }
    this.hooks.onSelectionChange()
  }

  // ====================== 命中检测 ======================

  private localXY(e: MouseEvent): { x: number; y: number } | null {
    const rect = this.els.renderArea.getBoundingClientRect()
    return { x: e.clientX - rect.left, y: e.clientY - rect.top }
  }

  private hitRegion(e: MouseEvent): Hit {
    const r = this.renderer
    const p = this.localXY(e)
    if (!r || !p) return { region: 'none' }
    const hw = r.metrics.rowHeaderWidth
    const hh = r.metrics.colHeaderHeight
    if (p.x < hw && p.y < hh) return { region: 'corner' }
    if (p.x < hw) return { region: 'row', row: r.rowAtScreen(this.view, p.y) }
    if (p.y < hh) return { region: 'col', col: r.colAtScreen(this.view, p.x) }
    const cell = r.cellAtScreen(this.view, p.x, p.y)
    return cell ? { region: 'cell', row: cell.row, col: cell.col } : { region: 'none' }
  }

  private nearColBorder(x: number, y: number): { col: number } | null {
    const r = this.renderer
    if (!r || y >= r.metrics.colHeaderHeight) return null
    const col = r.colAtScreen(this.view, x)
    if (col < 0) return null
    const rect = r.screenRectOfCell(this.view, 0, col)
    if (Math.abs(x - (rect.x + rect.w)) <= 4) return { col }
    if (Math.abs(x - rect.x) <= 4 && col > 0) return { col: col - 1 }
    return null
  }
  private nearRowBorder(x: number, y: number): { row: number } | null {
    const r = this.renderer
    if (!r || x >= r.metrics.rowHeaderWidth) return null
    const row = r.rowAtScreen(this.view, y)
    if (row < 0) return null
    const rect = r.screenRectOfCell(this.view, row, 0)
    if (Math.abs(y - (rect.y + rect.h)) <= 4) return { row }
    if (Math.abs(y - rect.y) <= 4 && row > 0) return { row: row - 1 }
    return null
  }

  // ====================== 鼠标交互 ======================

  onMouseDown(e: MouseEvent): void {
    if (e.button !== 0) return
    this.els.scroller.focus()
    const r = this.renderer
    const p = this.localXY(e)
    // 自动筛选下拉按钮(优先于一切)
    if (r && p) {
      const fcol = r.filterButtonAt(this.view, p.x, p.y)
      if (fcol != null) {
        this.openFilterPopup(fcol)
        return
      }
    }
    // 透视表行分组折叠/展开按钮(功能开启时)
    if (r && p && this.editCfg.pivotTable) {
      const tg = r.pivotToggleAt(this.view, p.x, p.y)
      if (tg) {
        this.togglePivotGroup(tg.tableIdx, tg.key)
        return
      }
    }
    // 图片拖拽移动(editable;命中浮动图 → 进入 image 拖拽,优先于选区)
    if (r && p && this.editCfg.editable && this.sheet) {
      const imgIdx = this.imageHitAt(p)
      if (imgIdx >= 0) {
        this.dragMode = 'image'
        this.imageDragIdx = imgIdx
        this.imageDragStartRect = anchorRect(r.metrics, this.sheet.images[imgIdx])
        this.imageDragStartMouse = { x: p.x, y: p.y }
        this.imageDragBefore = cloneImageAnchor(this.sheet.images[imgIdx])
        this.edit.ensureBaseline()
        return
      }
    }
    // 表头边界拖拽改宽高(优先于选择)
    if (r && p) {
      if (p.y < r.metrics.colHeaderHeight) {
        const b = this.nearColBorder(p.x, p.y)
        if (b) {
          this.dragMode = 'resize-col'
          this.resizeTarget = b.col
          this.resizeStartPos = p.x
          this.resizeStartSize = r.metrics.colWidth(b.col)
          this.beginResizeRecord('col', b.col)
          return
        }
      } else if (p.x < r.metrics.rowHeaderWidth) {
        const b = this.nearRowBorder(p.x, p.y)
        if (b) {
          this.dragMode = 'resize-row'
          this.resizeTarget = b.row
          this.resizeStartPos = p.y
          this.resizeStartSize = r.metrics.rowHeight(b.row)
          this.beginResizeRecord('row', b.row)
          return
        }
      }
    }
    const hit = this.hitRegion(e)
    this.dragMoved = false
    this.dragStartXY = this.localXY(e)
    if (hit.region === 'corner') {
      this.selectAll()
      this.dragMode = 'none'
    } else if (hit.region === 'row') {
      this.dragMode = 'row'
      this.setRows(hit.row, e.shiftKey)
    } else if (hit.region === 'col') {
      this.dragMode = 'col'
      this.setCols(hit.col, e.shiftKey)
    } else if (hit.region === 'cell') {
      this.dragMode = 'cell'
      this.setCell({ row: hit.row, col: hit.col }, e.shiftKey)
    } else {
      this.dragMode = 'none'
    }
    this.render()
  }

  onMouseMove(e: MouseEvent): void {
    if (this.dragMode !== 'none') {
      const r = this.renderer
      const p = this.localXY(e)
      if (!r || !p) return
      // 死区:移动超过 3px 才算"拖动"(避免单击时的微抖被当拖拽 → 误吞单击放大/单击语义)
      if (!this.dragStartXY || Math.abs(p.x - this.dragStartXY.x) > 3 || Math.abs(p.y - this.dragStartXY.y) > 3) {
        this.dragMoved = true
      }
      if (this.dragMode === 'image' && this.imageDragStartRect && this.sheet) {
        const dx = p.x - this.imageDragStartMouse.x
        const dy = p.y - this.imageDragStartMouse.y
        setImageRect(
          this.sheet,
          this.imageDragIdx,
          {
            left: this.imageDragStartRect.left + dx,
            top: this.imageDragStartRect.top + dy,
            width: this.imageDragStartRect.width,
            height: this.imageDragStartRect.height,
          },
          this.view.zoom,
        )
        this.render() // 重定位叠加层(图片 el 按 index 重读锚点)
        return
      }
      if (this.dragMode === 'resize-col') {
        this.setColWidthPx(this.resizeTarget, this.resizeStartSize + (p.x - this.resizeStartPos))
        return
      }
      if (this.dragMode === 'resize-row') {
        this.setRowHeightPx(this.resizeTarget, this.resizeStartSize + (p.y - this.resizeStartPos))
        return
      }
      if (this.dragMode === 'cell') {
        const cell = r.cellAtScreen(this.view, p.x, p.y)
        if (cell) {
          this.selActive = cell
          this.hooks.onSelectionChange()
          this.scheduleRender()
        }
      } else if (this.dragMode === 'row') {
        const row = r.rowAtScreen(this.view, p.y)
        if (row >= 0) {
          this.selActive = { row, col: 0 }
          this.hooks.onSelectionChange()
          this.scheduleRender()
        }
      } else {
        const col = r.colAtScreen(this.view, p.x)
        if (col >= 0) {
          this.selActive = { row: 0, col }
          this.hooks.onSelectionChange()
          this.scheduleRender()
        }
      }
      return
    }
    this.updateHover(e)
  }

  onMouseUp(e: MouseEvent): void {
    if (this.dragMode === 'cell' && !this.dragMoved) {
      const hit = this.hitRegion(e)
      const r = this.renderer
      if (hit.region === 'cell' && r) {
        this.hooks.onCellClick(hit.row, hit.col, r.cellText(hit.row, hit.col))
        const link = r.cellHyperlink(hit.row, hit.col)
        if (link) this.hooks.onHyperlink(link, { row: hit.row, col: hit.col })
        // 阅读模式单击图片 → 放大灯箱(编辑模式下单击仍只选中,放大走右键菜单)
        if (this.lightboxEnabled && !this.editCfg.editable && !this.isEditing()) {
          const p = this.localXY(e)
          const imgIdx = p ? this.imageHitAt(p) : -1 // 浮动图盖在格上,优先命中
          if (imgIdx >= 0) {
            const img = this.sheet?.images[imgIdx]
            if (img?.src) this.openImageLightbox(img.src, `image-${imgIdx + 1}.${img.mime?.split('/')[1] || 'png'}`, img.mime)
          } else {
            const ci = this.getCellImageAt(hit.row, hit.col) // 内嵌图(DISPIMG,画在 canvas)
            if (ci?.src) this.openImageLightbox(ci.src, `${ci.id}.${ci.mime?.split('/')[1] || 'png'}`, ci.mime)
          }
        }
      }
    } else if (this.dragMode === 'resize-col') {
      this.endResizeRecord('col', this.resizeTarget)
    } else if (this.dragMode === 'resize-row') {
      this.endResizeRecord('row', this.resizeTarget)
    } else if (this.dragMode === 'image') {
      if (this.dragMoved && this.imageDragBefore && this.sheet) {
        const after = cloneImageAnchor(this.sheet.images[this.imageDragIdx])
        this.edit.recordImageEdit(this.imageDragIdx, this.imageDragBefore, after)
      }
      this.imageDragBefore = null
      this.imageDragStartRect = null
      this.imageDragIdx = -1
    }
    this.dragMode = 'none'
  }

  /** 命中最上层浮动图(editable;p 为 render-area 坐标含表头);返回 index 或 -1。 */
  private imageHitAt(p: { x: number; y: number }): number {
    const r = this.renderer
    const s = this.sheet
    if (!r || !s) return -1
    const hw = r.metrics.rowHeaderWidth
    const hh = r.metrics.colHeaderHeight
    if (p.x < hw || p.y < hh) return -1 // 表头区不算
    for (let i = s.images.length - 1; i >= 0; i--) {
      const rect = anchorRect(r.metrics, s.images[i])
      const x = hw + rect.left - this.view.scrollX
      const y = hh + rect.top - this.view.scrollY
      if (p.x >= x && p.x <= x + rect.width && p.y >= y && p.y <= y + rect.height) return i
    }
    return -1
  }

  // ---- resize → 命令栈记账(仅 editable;E3.5) ----
  /** 拖拽起始:editable 时捕获 baseline + 起始维度信息(供 undo / dim-change)。 */
  private beginResizeRecord(axis: 'col' | 'row', index: number): void {
    if (!this.editCfg.editable || !this.sheet) return
    this.edit.ensureBaseline()
    if (axis === 'col') {
      const info = this.sheet.columns.get(index)
      this.resizeStartInfo = info ? { ...info } : null
      this.resizeStartModelSize = info?.width ?? this.sheet.defaultColWidth
    } else {
      const info = this.sheet.rows.get(index)
      this.resizeStartInfo = info ? { ...info } : null
      this.resizeStartModelSize = info?.height ?? this.sheet.defaultRowHeight
    }
  }
  /** 拖拽结束:模型已被 renderer 改完,补登一条 undo 项 + 发 dim-change。 */
  private endResizeRecord(axis: 'col' | 'row', index: number): void {
    if (!this.editCfg.editable || !this.sheet) return
    const after =
      axis === 'col'
        ? (this.sheet.columns.get(index)?.width ?? this.sheet.defaultColWidth)
        : (this.sheet.rows.get(index)?.height ?? this.sheet.defaultRowHeight)
    this.edit.recordDimEdit(axis, index, this.resizeStartInfo, this.resizeStartModelSize, after)
    this.resizeStartInfo = null
  }

  onMouseLeave(): void {
    this.hooks.onTooltip(null)
  }

  /** 右键上下文菜单(G3;仅 editable;只读用浏览器默认菜单)。壳把 contextmenu 事件转给它。 */
  /**
   * 右键事件入口(Plan C 全面开放):
   *  1. 算 ctx(选区 + 单/多格 + 活动格);右键不在当前选区 → 先选中(仿 Excel)
   *  2. 算内置 items(`buildBuiltinContextMenuItems`,editable 时才有,否则空)
   *  3. 跑用户 `setContextMenuTransform` 回调(可加/减/重排)
   *  4. 触发 `onContextMenuBefore` 钩子(可 preventDefault)
   *  5. 没被 prevent + items 非空 → 弹内置浮层;无论是否弹,都触发 `onContextMenuShow`(让用户也能自渲染)
   */
  onContextMenu(e: MouseEvent): void {
    const ctx = this.buildContextMenuCtx(e)
    if (!ctx) return
    let items = this.buildBuiltinContextMenuItems(ctx) // 只读模式也给(只含复制等不改数据的项),编辑模式给全套
    if (this.ctxMenuTransform) {
      const next = this.ctxMenuTransform(ctx, items)
      if (Array.isArray(next)) items = next
    }
    let prevented = false
    const preventDefault = () => { prevented = true }
    this.hooks.onContextMenuBefore?.({ event: e, ctx, items, preventDefault })
    if (prevented) {
      e.preventDefault()
      this.hooks.onContextMenuShow?.({ x: e.clientX, y: e.clientY, ctx, items })
      return
    }
    if (items.length) {
      e.preventDefault()
      this.menuHost.show(e.clientX, e.clientY, items)
    }
    this.hooks.onContextMenuShow?.({ x: e.clientX, y: e.clientY, ctx, items })
  }

  /** 算右键 ctx:命中格 + 选区调整;非内容区 / 无 renderer 返 null */
  private buildContextMenuCtx(e: MouseEvent): ContextMenuCtx | null {
    const r = this.renderer
    const p = this.localXY(e)
    if (!r || !p) return null
    const hit = r.cellAtScreen(this.view, p.x, p.y)
    let sel = this.getSelection()
    if (hit && (!sel || hit.row < sel.top || hit.row > sel.bottom || hit.col < sel.left || hit.col > sel.right)) {
      this.selectCell(hit.row, hit.col)
      this.render()
      sel = this.getSelection()
    }
    if (!sel || !this.sheet || !this.workbook) return null
    const range = { ...sel }
    const single = range.top === range.bottom && range.left === range.right
    return {
      range,
      single,
      activeCell: { row: range.top, col: range.left },
      sheet: this.sheet,
      workbook: this.workbook,
      editable: !!this.editCfg.editable,
    }
  }

  /** 内置右键菜单项(独立提取,便于 transform 回调拿到再二次加工)。
   *  只读模式:只返回**不改数据**的项(复制) —— 复制不修改数据源,只读也该能用;编辑模式给全套。 */
  buildBuiltinContextMenuItems(ctx: ContextMenuCtx): MenuItem[] {
    const range = ctx.range
    // 复制不改数据 → 任何模式都给
    const copyItem: MenuItem = { label: '复制', action: () => void this.copySelection() }
    if (!ctx.editable) return [copyItem]
    const rows = range.bottom - range.top + 1
    const cols = range.right - range.left + 1
    const single = ctx.single
    // Phase A 补漏 (2026-06-08): 各菜单项 disabled 反映 editable 闸门 (UX 跟权限一致)
    // anyEditable: 区域内至少 1 格可编辑 (粘贴/清除/换行/换需要)
    // allEditable: 区域内全部可编辑 (合并/拆分需要 — 否则 emit permission-denied 拒绝)
    let anyEditable = false
    let allEditable = true
    for (let r = range.top; r <= range.bottom; r++) {
      for (let c = range.left; c <= range.right; c++) {
        if (this.isCellEditable(r, c)) anyEditable = true
        else allEditable = false
      }
    }
    const items: MenuItem[] = [
      copyItem,
      { label: '粘贴', disabled: !anyEditable, action: () => void this.pasteFromClipboard() },
      {
        label: '选择性粘贴',
        disabled: !anyEditable,
        children: [
          { label: '覆盖格式(贴近源)', action: () => void this.pasteFromClipboard(DEFAULT_PASTE_BEHAVIOR) },
          { label: '保留原样式(仅值)', action: () => void this.pasteFromClipboard(PASTE_PRESET_VALUES_ONLY) },
        ],
      },
      { separator: true },
      { label: `在上方插入 ${rows} 行`, action: () => this.insertRows(range.top, rows) },
      { label: `在左侧插入 ${cols} 列`, action: () => this.insertCols(range.left, cols) },
      { label: `删除 ${rows} 行`, disabled: !allEditable, action: () => this.deleteRows(range.top, rows) },
      { label: `删除 ${cols} 列`, disabled: !allEditable, action: () => this.deleteCols(range.left, cols) },
      { separator: true },
      { label: '合并单元格', disabled: single || !allEditable, action: () => this.mergeCells(range) },
      { label: '拆分单元格', disabled: !allEditable, action: () => this.unmergeCells(range) },
      {
        label: (this.getSelectionWrapState() === 'all' ? '✓ ' : '') + '自动换行',
        disabled: !anyEditable,
        action: () => this.toggleWrapTextOnSelection(),
      },
      { separator: true },
      { label: '清除内容', disabled: !anyEditable, action: () => this.clearRange(range) },
    ]
    // 选区批量(多格)— 浮动图批量嵌入 / 嵌入图批量浮动化
    if (!single) {
      const imgs = ctx.sheet.images ?? []
      const floatsInRange = imgs.reduce((n, _, i) => {
        const a = this.imageCellOf(i)
        return n + (!!a && a.row >= range.top && a.row <= range.bottom && a.col >= range.left && a.col <= range.right ? 1 : 0)
      }, 0)
      let dispCount = 0
      for (let r = range.top; r <= range.bottom; r++) {
        for (let c = range.left; c <= range.right; c++) {
          if (ctx.sheet.cells.get(cellKey(r, c))?.dispImgId) dispCount++
        }
      }
      if (floatsInRange > 0 || dispCount > 0) {
        const rangeItems: MenuItem[] = []
        if (floatsInRange > 0) rangeItems.push({ label: `选区浮动图全部嵌入(${floatsInRange} 张)`, action: () => this.convertImagesInRangeToCell(range) })
        if (dispCount > 0) rangeItems.push({ label: `选区内嵌图全部浮动化(${dispCount} 张)`, action: () => this.convertCellImagesInRangeToFloat(range) })
        items.push({ separator: true }, ...rangeItems)
      }
    }
    // WPS 单元格内嵌图(DISPIMG)⇄ 浮动图互转(单格时才有意义)
    if (single) {
      const r = range.top
      const c = range.left
      const activeCell = ctx.sheet.cells.get(cellKey(r, c))
      const imgs = ctx.sheet.images ?? []
      const convertItems: MenuItem[] = []
      if (activeCell?.dispImgId) {
        const ci = this.getCellImageAt(r, c)
        if (this.lightboxEnabled && ci?.src) {
          convertItems.push({
            label: '查看大图 / 下载原图',
            action: () => this.openImageLightbox(ci.src, `${ci.id}.${ci.mime?.split('/')[1] || 'png'}`, ci.mime),
          })
        }
        convertItems.push({ label: '内嵌图转为浮动图', action: () => this.convertCellImageToFloat(r, c) })
      }
      if (imgs.length) {
        const hereIdx = imgs.findIndex((_, i) => {
          const a = this.imageCellOf(i)
          return !!a && a.row === r && a.col === c
        })
        if (hereIdx >= 0) convertItems.push({ label: '将此处浮动图嵌入单元格', action: () => this.convertImageToCellAuto(hereIdx) })
        const inCol = imgs.reduce((n, _, i) => n + (this.imageCellOf(i)?.col === c ? 1 : 0), 0)
        if (inCol > 0) convertItems.push({ label: `整列浮动图嵌入单元格(${inCol} 张)`, action: () => this.convertAllImagesToCells(c) })
        convertItems.push({ label: `整表浮动图嵌入单元格(${imgs.length} 张)`, action: () => this.convertAllImagesToCells() })
      }
      if (convertItems.length) items.push({ separator: true }, ...convertItems)
    }
    return items
  }

  /** 设置用户 transform 回调(`(ctx, items) => MenuItem[] | undefined`);壳侧调用 */
  setContextMenuTransform(fn: ContextMenuTransform | null | undefined): void {
    this.ctxMenuTransform = fn ?? null
  }
  /** 程序化打开菜单(键盘 Shift+F10 / 自定义触发);items 不给则按当前选区算内置 */
  openContextMenu(x: number, y: number, items?: MenuItem[]): void {
    if (items && items.length) {
      this.menuHost.show(x, y, items)
      return
    }
    const sel = this.getSelection()
    if (!sel || !this.sheet || !this.workbook) return
    const ctx: ContextMenuCtx = {
      range: { ...sel },
      single: sel.top === sel.bottom && sel.left === sel.right,
      activeCell: { row: sel.top, col: sel.left },
      sheet: this.sheet,
      workbook: this.workbook,
      editable: !!this.editCfg.editable,
    }
    let it = this.editCfg.editable ? this.buildBuiltinContextMenuItems(ctx) : []
    if (this.ctxMenuTransform) {
      const next = this.ctxMenuTransform(ctx, it)
      if (Array.isArray(next)) it = next
    }
    if (it.length) this.menuHost.show(x, y, it)
  }
  /** 关闭当前菜单(无打开则 no-op) */
  closeContextMenu(): void {
    this.menuHost.close()
  }

  private updateHover(e: MouseEvent): void {
    const r = this.renderer
    const sc = this.els.scroller
    const p = this.localXY(e)
    if (!r || !p) {
      this.hooks.onTooltip(null)
      return
    }
    // 表头边界 → 改宽高光标
    if (p.y < r.metrics.colHeaderHeight && this.nearColBorder(p.x, p.y)) {
      sc.style.cursor = 'col-resize'
      this.hooks.onTooltip(null)
      return
    }
    if (p.x < r.metrics.rowHeaderWidth && this.nearRowBorder(p.x, p.y)) {
      sc.style.cursor = 'row-resize'
      this.hooks.onTooltip(null)
      return
    }
    const cell = r.cellAtScreen(this.view, p.x, p.y)
    if (!cell) {
      this.hooks.onTooltip(null)
      sc.style.cursor = ''
      return
    }
    // Phase C 2026-06-08: editable=true 但该格只读 → not-allowed; 否则按超链接/默认
    if (this.editCfg.editable && !this.isCellEditable(cell.row, cell.col)) {
      sc.style.cursor = r.cellHyperlink(cell.row, cell.col) ? 'pointer' : 'not-allowed'
    } else {
      sc.style.cursor = r.cellHyperlink(cell.row, cell.col) ? 'pointer' : 'cell'
    }
    const tx = p.x + 14
    const ty = p.y + 18
    const comment = r.commentAt(cell.row, cell.col)
    if (comment) {
      this.hooks.onTooltip({ text: comment, x: tx, y: ty, kind: 'comment' })
      return
    }
    const full = r.overflowTextAt(cell.row, cell.col)
    this.hooks.onTooltip(full ? { text: full, x: tx, y: ty, kind: 'overflow' } : null)
  }

  onDblClick(e: MouseEvent): void {
    const r = this.renderer
    const p = this.localXY(e)
    if (!r || !p) return
    const colHit = this.nearColBorder(p.x, p.y)
    const rowHit = colHit ? null : this.nearRowBorder(p.x, p.y)
    if (colHit) {
      this.autoFitColumn(colHit.col)
    } else if (rowHit) {
      this.autoFitRow(rowHit.row)
    } else {
      const cell = r.cellAtScreen(this.view, p.x, p.y)
      if (!cell) return
      // 可编辑 → 双击进入编辑;否则发双击事件(向后兼容)
      if (this.editCfg.editable && this.isCellEditable(cell.row, cell.col)) this.beginEdit(cell.row, cell.col)
      else this.hooks.onCellDblClick(cell.row, cell.col, r.cellText(cell.row, cell.col))
    }
  }

  // ====================== 键盘交互 ======================

  private pageRows(): number {
    const r = this.renderer
    if (!r) return 10
    return Math.max(1, Math.floor((this.view.height - r.metrics.colHeaderHeight) / r.defaultRowPx) - 1)
  }

  onKeyDown(e: KeyboardEvent): void {
    // 编辑模式下的 撤销/重做(Ctrl+Z / Ctrl+Shift+Z / Ctrl+Y)
    if (this.editCfg.editable && (e.ctrlKey || e.metaKey) && (e.key === 'z' || e.key === 'Z')) {
      if (e.shiftKey) this.edit.redo()
      else this.edit.undo()
      e.preventDefault()
      return
    }
    if (this.editCfg.editable && (e.ctrlKey || e.metaKey) && (e.key === 'y' || e.key === 'Y')) {
      this.edit.redo()
      e.preventDefault()
      return
    }
    if ((e.ctrlKey || e.metaKey) && (e.key === 'c' || e.key === 'C')) {
      void this.copySelection()
      e.preventDefault()
      return
    }
    // Ctrl+V 不在此拦截:放行原生粘贴 → 触发 scroller 的 paste 事件(onPaste 拿原始 HTML);
    // 在此 preventDefault 会**阻止** paste 事件,反而退回净化路径。
    if ((e.ctrlKey || e.metaKey) && (e.key === 'a' || e.key === 'A')) {
      this.selectAll()
      this.render()
      e.preventDefault()
      return
    }
    // 编辑进入(可编辑活动格、当前未在编辑):F2 / 打字 / Delete-Backspace 清空
    if (this.editCfg.editable && this.selActive && !this.isEditing()) {
      const { row, col } = this.selActive
      if (e.key === 'F2') {
        this.beginEdit(row, col)
        e.preventDefault()
        return
      }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        if (this.isCellEditable(row, col)) this.edit.clearRange({ top: row, left: col, bottom: row, right: col })
        e.preventDefault()
        return
      }
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        if (this.beginEdit(row, col, e.key)) {
          e.preventDefault()
          return
        }
      }
    }
    const r = this.renderer
    if (!r || !this.selActive) return
    const maxRow = r.metrics.rows - 1
    const maxCol = r.metrics.cols - 1
    const ctrl = e.ctrlKey || e.metaKey
    let { row, col } = this.selActive
    let handled = true
    switch (e.key) {
      case 'ArrowUp': if (ctrl) { const j = jumpEdge(r, row, col, -1, 0); row = j.row; col = j.col } else row = Math.max(0, row - 1); break
      case 'ArrowDown': if (ctrl) { const j = jumpEdge(r, row, col, 1, 0); row = j.row; col = j.col } else row = Math.min(maxRow, row + 1); break
      case 'ArrowLeft': if (ctrl) { const j = jumpEdge(r, row, col, 0, -1); row = j.row; col = j.col } else col = Math.max(0, col - 1); break
      case 'ArrowRight': if (ctrl) { const j = jumpEdge(r, row, col, 0, 1); row = j.row; col = j.col } else col = Math.min(maxCol, col + 1); break
      case 'Home': col = 0; if (ctrl) row = 0; break
      case 'End': col = maxCol; if (e.ctrlKey) row = maxRow; break
      case 'PageUp': row = Math.max(0, row - this.pageRows()); break
      case 'PageDown': row = Math.min(maxRow, row + this.pageRows()); break
      case 'Enter': row = Math.min(maxRow, row + 1); break
      case 'Tab': col = e.shiftKey ? Math.max(0, col - 1) : Math.min(maxCol, col + 1); break
      default: handled = false
    }
    if (!handled) return
    e.preventDefault()
    const m = r.mergeAt(row, col)
    if (m) {
      row = m.top
      col = m.left
    }
    this.selMode = 'range'
    this.selActive = { row, col }
    const extend = e.shiftKey && e.key !== 'Tab'
    if (!extend) this.selAnchor = { row, col }
    this.scrollActiveIntoView()
    this.hooks.onSelectionChange()
    this.render()
  }

  private scrollActiveIntoView(): void {
    if (this.selActive) this.scrollCellIntoView(this.selActive)
  }

  private scrollCellIntoView(c: Cell): void {
    const r = this.renderer
    const sc = this.els.scroller
    if (!r || !c) return
    const hw = r.metrics.rowHeaderWidth
    const hh = r.metrics.colHeaderHeight
    const fz = r.freezeGeometry
    let sx = sc.scrollLeft
    let sy = sc.scrollTop
    if (c.col >= fz.frozenCols) {
      const cl = r.metrics.colLeft(c.col)
      const cr = cl + r.metrics.colWidth(c.col)
      const viewW = this.view.width - hw
      if (cr > sx + viewW) sx = cr - viewW
      if (cl < sx + fz.frozenWidth) sx = cl - fz.frozenWidth
    }
    if (c.row >= fz.frozenRows) {
      const ct = r.metrics.rowTop(c.row)
      const cb = ct + r.metrics.rowHeight(c.row)
      const viewH = this.view.height - hh
      if (cb > sy + viewH) sy = cb - viewH
      if (ct < sy + fz.frozenHeight) sy = ct - fz.frozenHeight
    }
    sx = Math.max(0, sx)
    sy = Math.max(0, sy)
    if (sx !== sc.scrollLeft || sy !== sc.scrollTop) {
      sc.scrollLeft = sx
      sc.scrollTop = sy
      this.view.scrollX = sx
      this.view.scrollY = sy
    }
  }

  // ====================== 复制 ======================

  async copySelection(): Promise<void> {
    const r = this.renderer
    const s = this.getSelection()
    const sheet = this.sheet
    const wb = this.workbook
    if (!r || !s || !sheet || !wb) return
    // 防超大选区卡死: 复制范围软上限
    const rowEnd = Math.min(s.bottom, s.top + 4999)
    const colEnd = Math.min(s.right, s.left + 255)
    const range = { top: s.top, left: s.left, bottom: rowEnd, right: colEnd }

    // 合并区(完全落在复制区内):左上格出 colspan/rowspan,其余被覆盖格不出 <td>
    const covered = new Set<string>()
    const spanAt = new Map<string, { rs: number; cs: number }>()
    for (const m of sheet.merges) {
      if (m.top < range.top || m.left < range.left || m.bottom > range.bottom || m.right > range.right) continue
      spanAt.set(cellKey(m.top, m.left), { rs: m.bottom - m.top + 1, cs: m.right - m.left + 1 })
      for (let rr = m.top; rr <= m.bottom; rr++) for (let cc = m.left; cc <= m.right; cc++) if (rr !== m.top || cc !== m.left) covered.add(cellKey(rr, cc))
    }
    // 先轻量统计图片总字节(不编码),判断是否超预算 → 降级为"无图 1:1 复制"
    let imageBytes = 0
    for (let row = range.top; row <= range.bottom; row++) {
      for (let col = range.left; col <= range.right; col++) {
        const id = sheet.cells.get(cellKey(row, col))?.dispImgId
        const ci = id ? wb.cellImages?.get(id) : undefined
        if (ci?.bytes) imageBytes += ci.bytes.length
      }
    }
    for (const im of sheet.images) {
      if (im.from.row >= range.top && im.from.row <= range.bottom && im.from.col >= range.left && im.from.col <= range.right && im.bytes) imageBytes += im.bytes.length
    }
    const dropImages = imageBytes > CLIP_IMAGE_BUDGET_BYTES
    if (dropImages) {
      this.hooks.onEditEvent('permission-denied', {
        reason: 'copy',
        cells: [],
        message: `复制内容含图过多(约 ${Math.round(imageBytes / 1024 / 1024)} MB),已按无图复制以避免剪贴板超限`,
      })
    }

    // 图片按格归位(降级时跳过)。key 与快照引用对齐:DISPIMG → c:id,浮动 → f:序号(同 serializeSnapshot 顺序)
    // 带上 w/h(逻辑 px):WPS/Excel 解析剪贴板 HTML 用 <img width/height> 属性,不认 CSS max-width —— 不给就按原图
    // 像素尺寸贴入(产品图常几百上千 px → 贴进去巨大)。DISPIMG 按所在格大小,浮动图按其 EMU 尺寸。
    const colPx = (col: number) => Math.round(sheet.columns.get(col)?.width ?? sheet.defaultColWidth)
    const rowPx = (row: number) => Math.round(sheet.rows.get(row)?.height ?? sheet.defaultRowHeight)
    const imgAt = new Map<string, Array<{ key: string; url: string; w: number; h: number }>>()
    if (!dropImages) {
      const push = (k: string, key: string, url: string, w: number, h: number) => {
        const a = imgAt.get(k) ?? []
        a.push({ key, url, w, h })
        imgAt.set(k, a)
      }
      for (let row = range.top; row <= range.bottom; row++) {
        for (let col = range.left; col <= range.right; col++) {
          const id = sheet.cells.get(cellKey(row, col))?.dispImgId
          const ci = id ? wb.cellImages?.get(id) : undefined
          if (id && ci?.bytes && ci.mime) push(cellKey(row, col), `c:${id}`, `data:${ci.mime};base64,${bytesToB64(ci.bytes)}`, colPx(col), rowPx(row))
        }
      }
      let fi = 0
      for (const im of sheet.images) {
        if (im.from.row < range.top || im.from.row > range.bottom || im.from.col < range.left || im.from.col > range.right || !im.bytes || !im.mime) continue
        const w = im.extWidthEmu ? Math.round(emuToPx(im.extWidthEmu)) : colPx(im.from.col)
        const h = im.extHeightEmu ? Math.round(emuToPx(im.extHeightEmu)) : rowPx(im.from.row)
        push(cellKey(im.from.row, im.from.col), `f:${fi}`, `data:${im.mime};base64,${bytesToB64(im.bytes)}`, w, h)
        fi++
      }
    }

    const lines: string[] = []
    const htmlRows: string[] = []
    for (let row = range.top; row <= range.bottom; row++) {
      const cells: string[] = []
      const htmlCells: string[] = []
      for (let col = range.left; col <= range.right; col++) {
        const text = r.cellText(row, col)
        cells.push(text)
        if (covered.has(cellKey(row, col))) continue // 被合并覆盖,不出 <td>(TSV 仍占位保持列对齐)
        const css = r.cellInlineStyle(row, col)
        const span = spanAt.get(cellKey(row, col))
        const spanAttr = span ? `${span.rs > 1 ? ` rowspan="${span.rs}"` : ''}${span.cs > 1 ? ` colspan="${span.cs}"` : ''}` : ''
        const imgHtml = (imgAt.get(cellKey(row, col)) ?? []).map((g) => `<img data-clip-img="${g.key}" width="${g.w}" height="${g.h}" src="${g.url}" style="width:${g.w}px;height:${g.h}px" />`).join('')
        htmlCells.push(`<td${spanAttr}${css ? ` style="${css}"` : ''}>${escapeHtml(text)}${imgHtml}</td>`)
      }
      lines.push(cells.join('\t'))
      htmlRows.push(`<tr>${htmlCells.join('')}</tr>`)
    }
    const tsv = lines.join('\n')
    // 瘦身快照嵌进 data-ooxml-clip:图片字节不进快照(只引用),由可见 <img> 携带 → 避免双重 base64;
    // 降级时去掉图片。粘贴时本组件读快照 + 回填 <img> 字节做 1:1;外部应用忽略属性,只读可见 table。
    const snap = serializeSnapshot(sheet, wb, range, { withImageBytes: false })
    const clip = encodeSnapshot(dropImages ? withoutImages(snap) : snap)
    const html = `<table data-ooxml-clip="${clip}" border="1" style="border-collapse:collapse">${htmlRows.join('')}</table>`
    try {
      // 优先写 text/plain + text/html(粘到 Word/Excel 保留表格与格式)
      const ClipItem = (window as unknown as { ClipboardItem?: typeof ClipboardItem }).ClipboardItem
      if (ClipItem && navigator.clipboard?.write) {
        await navigator.clipboard.write([
          new ClipItem({
            'text/plain': new Blob([tsv], { type: 'text/plain' }),
            'text/html': new Blob([html], { type: 'text/html' }),
          }),
        ])
      } else {
        await navigator.clipboard.writeText(tsv)
      }
    } catch {
      /* 某些环境无剪贴板权限，静默忽略 */
    }
  }

  /**
   * paste 事件粘贴(Ctrl+V 主路径,绑在 scroller 上)。`e.clipboardData` 给的是**原始未净化** HTML/图片/文本 ——
   * 关键:`navigator.clipboard.read()`(pasteFromClipboard 用)会净化 HTML、删掉 `<style>` 块和注释,而 WPS/Excel
   * 的格式(CSS 类)、数字格式(mso-number-format)、内嵌图(VML o:gfxdata)全在那里面 → 走 read 必丢。paste 事件不净化。
   * 我们自己复制的快照(data-ooxml-clip)也一并原样拿到,1:1 不受影响。
   */
  /**
   * 统一出口(核心层唯一只读反馈点):**任何**改数据的操作撞只读 → EditController/控制器都经 `host.emit` 走到这里,
   * 按 `readOnlyPrompt` 配置(dialog/toast/none)弹一次内置提醒,再转壳事件。新增输入方式只要照常走 EditController
   * 的 API(其内 `isEditable` 中央闸门已逐格拦截 + emit permission-denied),无需各自重写只读检查/提醒。
   */
  private emitEditEvent(event: EditEventName, payload: unknown): void {
    if (event === 'permission-denied') {
      const p = payload as PermissionDeniedPayload
      // 改数据源的操作(paste/merge/unmerge/image-convert)撞只读 → 统一弹提醒;dimension(列宽行高=布局,非数据)不弹
      if (p?.reason && p.reason !== 'dimension') {
        this.readonlyPrompt.show(this.editCfg.readOnlyPrompt ?? 'dialog', p.message ?? '目标只读,操作已跳过', p.cells ?? [])
      }
    }
    this.hooks.onEditEvent(event, payload)
  }

  onPaste(e: ClipboardEvent): void {
    const dt = e.clipboardData
    if (!dt) return
    const html = dt.getData('text/html')
    const text = dt.getData('text/plain')
    const file = Array.from(dt.items ?? []).find((it) => it.kind === 'file' && it.type.startsWith('image/'))?.getAsFile()
    const hasContent = !!html || !!text || !!file
    // 只读提示(用户诉求):只读模式 / 落点只读 → 不静默,发 permission-denied 让壳 toast,用户知道为啥没粘上
    if (!this.editCfg.editable) {
      if (hasContent) { e.preventDefault(); this.emitEditEvent('permission-denied', { reason: 'paste', cells: [], message: '当前为只读模式,无法粘贴(开启编辑后再试)' } satisfies PermissionDeniedPayload) }
      return
    }
    // 落点/区域是否只读不在此预判 —— 交给 pasteRich/pasteSnapshot 的中央逐格 isEditable 闸门:可编辑的格照常粘,
    // 只读格自动跳过并收集 → 经 emitEditEvent 统一提醒。避免在输入层重写一遍只读检查(换输入方式不必重写)。
    if (html && this.pasteRichHtml(html)) { e.preventDefault(); return }
    if (file) { void this.pasteImageBlob(file); e.preventDefault(); return }
    if (text && this.pasteText(text)) e.preventDefault()
  }

  /**
   * 从系统剪贴板粘贴(右键菜单"粘贴"等无 paste 事件的入口用)。`clipboard.read()` 拿 text/html —— 注意它会**净化**
   * HTML(删 `<style>`/注释),所以从 WPS/Excel 粘的格式不如 Ctrl+V(走 onPaste 的原始 HTML)全。否则单图 / TSV 兜底。
   */
  async pasteFromClipboard(behaviorOverride?: Partial<PasteBehavior> | null): Promise<boolean> {
    if (!this.editCfg.editable) {
      this.emitEditEvent('permission-denied', { reason: 'paste', cells: [], message: '当前为只读模式,无法粘贴(开启编辑后再试)' } satisfies PermissionDeniedPayload)
      return false
    }
    type Clip = {
      read?: () => Promise<Array<{ types: string[]; getType: (t: string) => Promise<Blob> }>>
      readText?: () => Promise<string>
    }
    const clip = (navigator as unknown as { clipboard?: Clip }).clipboard
    try {
      if (clip?.read) {
        const items = await clip.read()
        for (const it of items) {
          if (it.types.includes('text/html')) {
            const html = await (await it.getType('text/html')).text()
            if (this.pasteRichHtml(html, undefined, behaviorOverride)) return true
          }
        }
        for (const it of items) {
          const imgType = it.types.find((t) => t.startsWith('image/'))
          if (imgType && (await this.pasteImageBlob(await it.getType(imgType)))) return true
        }
      }
    } catch {
      /* read 受限/无权限 → 回退 readText */
    }
    try {
      const text = await clip?.readText?.()
      return text ? this.pasteText(text) : false
    } catch {
      return false
    }
  }

  /**
   * 解析 Excel/WPS 复制的剪贴板 HTML(text/html)→ 富粘贴:值 + 字体/颜色/填充/边框/对齐 + 合并 + data-uri 图,
   * **整体单次撤销**。无 `<table>` 返 false(调用方回退 TSV)。at 缺省用活动格。
   */
  pasteRichHtml(html: string, at?: { row: number; col: number }, behaviorOverride?: Partial<PasteBehavior> | null): boolean {
    if (!this.editCfg.editable || !this.sheet) return false
    const sel = this.getSelection()
    const start = at ?? this.selActive ?? (sel ? { row: sel.top, col: sel.left } : null)
    if (!start) return false
    // 本组件自己复制的 → 用嵌在剪贴板里的完整快照做 1:1(跨实例 Vue3/Vue2/React 通用);否则走外部 HTML 近似解析
    const snap = parseSnapshotHtml(html)
    if (snap) return this.edit.pasteSnapshot(start, snap, behaviorOverride)
    const parsed = parseClipboardHtml(html)
    if (!parsed) return false
    return this.edit.pasteRich(start, parsed, behaviorOverride)
  }

  /** 读当前粘贴行为配置(完整)。 */
  getPasteBehavior(): PasteBehavior {
    return this.edit.getPasteBehavior()
  }
  /** 设粘贴行为默认(缺项回落默认);影响 Ctrl+V / 右键「粘贴」。右键「选择性粘贴」逐次预设走 pasteRichHtml 的第 3 参。 */
  setPasteBehavior(cfg: Partial<PasteBehavior> | null): void {
    this.edit.setPasteBehavior(cfg)
  }
  /** 打开「粘贴行为配置」面板(框架无关 DOM,三壳共用);应用即 setPasteBehavior。需 editable。 */
  openPasteConfigDialog(): boolean {
    if (!this.editCfg.editable) return false
    this.pasteConfigDialog.show({ current: this.edit.getPasteBehavior(), onSubmit: (cfg) => this.edit.setPasteBehavior(cfg) })
    return true
  }

  /** 把一张图片 blob 落到活动格(转内嵌图);剪贴板单图粘贴 / 拖文件进网格用。 */
  async pasteImageBlob(blob: Blob, at?: { row: number; col: number }): Promise<boolean> {
    if (!this.editCfg.editable || !this.sheet) return false
    const sel = this.getSelection()
    const start = at ?? this.selActive ?? (sel ? { row: sel.top, col: sel.left } : { row: 0, col: 0 })
    const bytes = new Uint8Array(await blob.arrayBuffer())
    const mime = blob.type || 'image/png'
    const idx = this.addImage({ src: '', bytes, mime, from: { col: start.col, row: start.row, colOffEmu: 0, rowOffEmu: 0 }, extWidthEmu: pxToEmu(96), extHeightEmu: pxToEmu(96) })
    return idx >= 0 ? this.convertImageToCell(idx, start.row, start.col) : false
  }

  /**
   * 把 TSV(Excel/表格复制的制表符分隔文本)粘到选区左上角(无 at 时用活动格)。
   * 值类型自动推断(纯数字串→数字、`=`→公式)、跳过只读格;入命令栈可撤销。返回是否有改动。
   */
  pasteText(text: string, at?: { row: number; col: number }): boolean {
    if (!this.editCfg.editable || !this.sheet) return false
    const grid = parseClipboardGrid(text)
    if (!grid.length) return false
    const sel = this.getSelection()
    const start = at ?? this.selActive ?? (sel ? { row: sel.top, col: sel.left } : null)
    if (!start) return false
    const top = start.row
    const left = start.col
    const bottom = top + grid.length - 1
    const right = left + Math.max(1, ...grid.map((r) => r.length)) - 1
    return this.edit.editRange({ top, left, bottom, right }, grid)
  }

  // ====================== 查找 ======================

  /** 查找状态快照(壳渲染 FindBar) */
  getFindState(): FindState {
    return {
      query: this.findQuery,
      matchCase: this.findMatchCase,
      wholeCell: this.findWholeCell,
      count: this.findHits.length,
      index: this.findIndex,
    }
  }

  setFindQuery(q: string): void {
    this.findQuery = q
    this.recomputeFind()
  }
  setFindMatchCase(b: boolean): void {
    this.findMatchCase = b
    this.recomputeFind()
  }
  setFindWholeCell(b: boolean): void {
    this.findWholeCell = b
    this.recomputeFind()
  }

  /** 重算命中并应用(query/选项变化时) */
  private recomputeFind(): void {
    const r = this.renderer
    if (!r || !this.findQuery) {
      this.findHits = []
      this.findIndex = -1
      r?.setFind([], -1)
      this.hooks.onFindChange()
      this.render()
      return
    }
    this.findHits = r.searchCells(this.findQuery, { matchCase: this.findMatchCase, wholeCell: this.findWholeCell })
    this.findIndex = this.findHits.length ? 0 : -1
    this.applyFind()
  }

  /** 把当前命中应用到渲染器 + 移动选区/滚动到视图 */
  private applyFind(): void {
    const r = this.renderer
    if (!r) return
    r.setFind(this.findHits, this.findIndex)
    this.hooks.onFindChange()
    const hit = this.findHits[this.findIndex]
    if (hit) this.selectCell(hit.row, hit.col) // 移动选区 + 滚动到视图 + 重绘
    else this.render()
  }

  findNext(): void {
    if (!this.findHits.length) return
    this.findIndex = (this.findIndex + 1) % this.findHits.length
    this.applyFind()
  }
  findPrev(): void {
    if (!this.findHits.length) return
    this.findIndex = (this.findIndex - 1 + this.findHits.length) % this.findHits.length
    this.applyFind()
  }

  /** 关闭查找: 清 query + 命中(保留 matchCase/wholeCell 选项),清除高亮 */
  clearFind(): void {
    this.findQuery = ''
    this.findHits = []
    this.findIndex = -1
    this.renderer?.setFind([], -1)
    this.hooks.onFindChange()
    this.render()
  }

  // ====================== 自动筛选 ======================

  /** 当前是否有列处于筛选中(工具栏「清除筛选」启用判断) */
  hasFilters(): boolean {
    return this.filterState.size > 0
  }

  /** 当前筛选浮层(壳渲染 FilterPopup) */
  getFilterPopup(): FilterPopupState | null {
    return this.filterPopup
  }

  /** 筛选数据区底行: 正常用 af.bottom;若 af 只含表头(bottom===top)则延伸到数据末行 */
  private filterDataBottom(): number {
    const s = this.sheet!
    const af = s.autoFilterRange!
    return af.bottom > af.top ? af.bottom : s.dimension.rows - 1
  }

  /** 某列(自动筛选数据区)的去重值,数值/中文自然排序 */
  private distinctColumnValues(col: number): string[] {
    const r = this.renderer
    const s = this.sheet
    if (!r || !s?.autoFilterRange) return []
    const af = s.autoFilterRange
    const set = new Set<string>()
    for (let row = af.top + 1; row <= this.filterDataBottom(); row++) set.add(r.cellText(row, col) || BLANK)
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
  }

  /** 重算筛选导致的隐藏行并应用到模型(行隐藏机制 → 几何归零) */
  private applyFilters(): void {
    const r = this.renderer
    const s = this.sheet
    if (!r || !s?.autoFilterRange) return
    const af = s.autoFilterRange
    const bottom = this.filterDataBottom()
    if (!this.filterOrigHidden.size) {
      for (let row = af.top + 1; row <= bottom; row++) this.filterOrigHidden.set(row, s.rows.get(row)?.hidden ?? false)
    }
    for (let row = af.top + 1; row <= bottom; row++) {
      const orig = this.filterOrigHidden.get(row) ?? false
      let excluded = false
      for (const [col, allowed] of this.filterState) {
        if (!allowed.has(r.cellText(row, col) || BLANK)) {
          excluded = true
          break
        }
      }
      const hidden = orig || excluded
      const info = s.rows.get(row)
      if (info) info.hidden = hidden
      else if (hidden) s.rows.set(row, { height: s.defaultRowHeight, hidden: true })
    }
    r.setFilteredCols(new Set(this.filterState.keys()))
    r.rebuildMetrics()
    this.refreshContentSize()
    this.clearSelection()
    this.hooks.onFilterChange()
    this.render()
  }

  /** 点中下拉按钮 / 命令式: 打开某列筛选浮层(算去重值 + 已选 + 屏幕位置) */
  openFilterPopup(col: number): void {
    const r = this.renderer
    const s = this.sheet
    if (!r || !s?.autoFilterRange) return
    const rect = r.cellScreenRect(this.view, s.autoFilterRange.top, col)
    let x = rect.x
    let y = rect.y + rect.h
    if (x + 228 > this.view.width) x = Math.max(0, this.view.width - 232)
    if (y + 320 > this.view.height) y = Math.max(0, rect.y - 320)
    this.filterPopup = {
      col,
      values: this.distinctColumnValues(col),
      selected: this.filterState.has(col) ? [...this.filterState.get(col)!] : [],
      x,
      y,
      sortDir: this.sortCol === col ? this.sortDir : null,
    }
    this.hooks.onFilterChange()
  }

  closeFilterPopup(): void {
    if (!this.filterPopup) return
    this.filterPopup = null
    this.hooks.onFilterChange()
  }

  /** 浮层「确定」: 勾选值落筛选态(全选=取消该列;子集=保留;空集=全隐藏)。 */
  applyFilterSelection(checked: string[]): void {
    const pop = this.filterPopup
    if (!pop) return
    const all = this.distinctColumnValues(pop.col)
    if (checked.length >= all.length) this.filterState.delete(pop.col)
    else this.filterState.set(pop.col, new Set(checked))
    this.filterPopup = null
    this.applyFilters()
  }

  /** 浮层「清除」: 取消该列筛选 */
  clearFilterColumn(): void {
    const pop = this.filterPopup
    if (!pop) return
    this.filterState.delete(pop.col)
    this.filterPopup = null
    this.applyFilters()
  }

  /** 清除当前表全部筛选 */
  clearAllFilters(): void {
    if (!this.filterState.size) return
    this.filterState.clear()
    this.applyFilters()
  }

  /**
   * 离开某表 / 取消自动筛选: 恢复被筛选隐藏的行,清空筛选态。
   * 传入要恢复的 sheet(切表时是"旧表",取消筛选时是当前表)。
   */
  resetFilter(sheet: SheetModel | null | undefined): void {
    if (sheet && this.filterOrigHidden.size) {
      for (const [row, orig] of this.filterOrigHidden) {
        const info = sheet.rows.get(row)
        if (info) info.hidden = orig
      }
    }
    this.filterOrigHidden.clear()
    this.filterState.clear()
    this.filterPopup = null
    this.hooks.onFilterChange()
  }

  /** 换工作簿: 仅清筛选态(旧 sheet 已废,不需恢复行) */
  clearFilterState(): void {
    this.filterOrigHidden.clear()
    this.filterState.clear()
    this.filterPopup = null
    this.hooks.onFilterChange()
  }

  /** 工具栏「筛选」: 切换自动筛选。无则按选区(或整张已用区)新建,使下拉按钮出现。 */
  toggleAutoFilter(): void {
    const s = this.sheet
    const r = this.renderer
    if (!s || !r) return
    if (s.autoFilterRange) {
      this.resetFilter(s) // 恢复筛选隐藏的行 + 清状态
      s.autoFilterRange = undefined
    } else {
      const sel = this.getSelection()
      const multi = sel && !(sel.top === sel.bottom && sel.left === sel.right)
      s.autoFilterRange = multi
        ? { ...sel! }
        : { top: 0, left: 0, bottom: Math.max(0, s.dimension.rows - 1), right: Math.max(0, s.dimension.cols - 1) }
    }
    r.setFilteredCols(new Set())
    r.rebuildMetrics()
    this.refreshContentSize()
    this.hooks.onFilterChange()
    this.render()
  }

  // ====================== 排序 ======================

  /** 当前排序状态(壳可读以显示指示) */
  getSortState(): { col: number; dir: 'asc' | 'desc' | null } {
    return { col: this.sortCol, dir: this.sortDir }
  }

  /** 按活动单元格所在列排序;若当前表未开启自动筛选,先按选区/已用区建立筛选范围。 */
  sortActiveColumn(dir: 'asc' | 'desc'): boolean {
    const s = this.sheet
    const active = this.getActiveCell()
    if (!s || !active) return false
    if (!s.autoFilterRange) this.toggleAutoFilter()
    if (!s.autoFilterRange || active.col < s.autoFilterRange.left || active.col > s.autoFilterRange.right) return false
    this.sortColumn(active.col, dir)
    return this.sortCol === active.col && this.sortDir === dir
  }

  /** 打开 WPS 式透视表入口:确认选区字段后生成静态透视汇总表。需 `pivotTable` + `editable` 配置。 */
  openPivotTableDialog(): boolean {
    const sel = this.getSelection()
    if (!this.editCfg.pivotTable) return false
    if (!sel || !this.renderer || !this.sheet || !this.editCfg.editable) return false
    const fields = this.pivotFieldOptions(sel)
    if (fields.length < 2) return false
    this.pivotDialog.show({
      rangeLabel: `${colLabel(sel.left)}${sel.top + 1}:${colLabel(sel.right)}${sel.bottom + 1}`,
      defaultOutputCell: `${colLabel(Math.min(sel.right + 2, MAX_GRID_COLS - 2))}${sel.top + 1}`,
      onSubmit: (output) => this.createPivotTableFromSelection({ output }),
    })
    return true
  }

  /** 兼容旧 API:基于当前选区创建静态透视汇总表。 */
  createPivotTableFromSelection(opts: { rowFieldIndex?: number; valueFieldIndex?: number; output?: PivotOutputChoice } = {}): boolean {
    const sel = this.getSelection()
    if (!sel) return this.failPivot('透视表创建失败:请先选择包含表头且至少两列两行的数据区域。')
    const layout: Partial<PivotTableLayout> = {}
    if (opts.rowFieldIndex != null) layout.rows = [opts.rowFieldIndex]
    if (opts.valueFieldIndex != null) layout.values = [{ field: opts.valueFieldIndex, summary: 'sum' }]
    return this.createPivotTable({ sourceRange: sel, output: opts.output, layout, showPanel: true })
  }

  /** 通过 API 直接创建静态透视表,不依赖页面选区或对话框。需 `pivotTable` + `editable` 配置。 */
  createPivotTable(opts: CreatePivotTableOptions): boolean {
    const r = this.renderer
    const wb = this.workbook
    if (!this.editCfg.pivotTable) return this.failPivot('透视表创建失败:未开启透视表功能(pivotTable 配置,默认关闭)。')
    if (!r || !wb) return this.failPivot('透视表创建失败:工作簿尚未加载完成。')
    if (!this.editCfg.editable) return this.failPivot('透视表创建失败:请先开启编辑模式。')
    const sourceIndex = opts.sourceSheetIndex ?? this.activeIndex
    const s = wb.sheets[sourceIndex]
    if (!s) return this.failPivot('透视表创建失败:源工作表不存在。')
    const sel = opts.sourceRange ?? this.getSelection()
    if (!sel || sel.bottom <= sel.top || sel.right <= sel.left) return this.failPivot('透视表创建失败:请先选择包含表头且至少两列两行的数据区域。')

    const headers: string[] = []
    for (let col = sel.left; col <= sel.right; col++) headers.push(this.pivotCellText(s, sel.top, col) || `字段${col - sel.left + 1}`)
    const fieldOptions = this.pivotFieldOptionsForSheet(s, sel)

    // 显式给了行或值字段 → 沿用(编程 API,缺一边自动补);否则空白起步(对话框),不猜字段,等用户在面板勾选。
    const explicit = !!(opts.layout?.rows?.length || opts.layout?.values?.length)
    let initialLayout: PivotTableLayout
    if (explicit) {
      const fieldCount = headers.length
      let rowFieldOffset = (opts.layout?.rows?.[0] ?? -1) - sel.left
      let valueFieldOffset = (opts.layout?.values?.[0]?.field ?? -1) - sel.left
      if (valueFieldOffset < 0) valueFieldOffset = (fieldOptions.find((f) => f.numeric)?.index ?? -1) - sel.left
      if (rowFieldOffset < 0) rowFieldOffset = (fieldOptions.find((f) => f.index !== sel.left + valueFieldOffset)?.index ?? -1) - sel.left
      if (valueFieldOffset < 0 || valueFieldOffset >= fieldCount) return this.failPivot('透视表创建失败:请选择有效的值字段。')
      if (rowFieldOffset < 0 || rowFieldOffset === valueFieldOffset) rowFieldOffset = valueFieldOffset === 0 ? 1 : 0
      if (rowFieldOffset < 0 || rowFieldOffset >= fieldCount) return this.failPivot('透视表创建失败:请选择有效的行字段。')
      initialLayout = normalizePivotLayout(opts.layout, sel.left + rowFieldOffset, sel.left + valueFieldOffset)
    } else {
      initialLayout = { filters: opts.layout?.filters?.map((rule) => ({ ...rule, values: rule.values?.slice() })) ?? [], columns: [], rows: [], values: [] }
    }
    const built = this.buildPivotRows(s, sel, fieldOptions, initialLayout, [])
    if (!built) return this.failPivot('透视表创建失败:值字段没有可汇总的数据。')

    const output = opts.output ?? { kind: 'current-sheet' as const, cell: `${colLabel(sel.right + 2)}${sel.top + 1}` }
    const at = output.kind === 'new-sheet' ? { row: 0, col: 0 } : parseCellRef(output.cell)
    if (!at) return this.failPivot('透视表创建失败:请输入有效的生成位置,例如 H1。')
    const outTop = at.row
    const outLeft = at.col
    const outWidth = Math.max(1, ...built.rows.map((row) => row.length))
    if (outLeft + outWidth - 1 >= MAX_GRID_COLS) return this.failPivot('透视表创建失败:生成位置超出最大列数。')
    const range = { top: outTop, left: outLeft, bottom: outTop + built.rows.length - 1, right: outLeft + outWidth - 1 }
    if (range.bottom >= MAX_GRID_ROWS) return this.failPivot('透视表创建失败:生成位置超出最大行数。')
    if (output.kind === 'current-sheet') {
      for (let row = range.top; row <= range.bottom; row++) {
        for (let col = range.left; col <= range.right; col++) {
          if (!this.isCellEditable(row, col)) return this.failPivot('透视表创建失败:生成区域包含只读单元格,请换一个位置或选择新建工作表。')
        }
      }
    }

    this.edit.ensureBaseline()
    const snap = cloneWorkbook(wb)
    const outSheet = output.kind === 'new-sheet' ? createPivotSheet(wb) : s
    for (let rr = 0; rr < built.rows.length; rr++) {
      for (let cc = 0; cc < built.rows[rr].length; cc++) setCellValue(outSheet, outTop + rr, outLeft + cc, built.rows[rr][cc])
    }
    const pivot: PivotTableModel = {
      name: `PivotTable${outSheet.pivotTables.length + 1}`,
      range,
      fields: headers,
      source: { sheetIndex: sourceIndex, range: { ...sel } },
      layout: clonePivotLayout(initialLayout),
      collapsed: [],
      rowGroups: built.groups.map((g) => ({ row: outTop + g.rowOffset, key: g.key })),
      buttons: [
        { row: outTop, col: outLeft, label: built.rowLabel, kind: 'row' as const },
        { row: outTop, col: outLeft + 1, label: built.valueLabel, kind: 'data' as const },
      ],
    }
    outSheet.pivotTables.push(pivot)
    this.edit.pushUndoExternal({ kind: 'restore-wb', snapshot: snap })
    this.edit.markDirtyExternal()
    if (output.kind === 'new-sheet') {
      wb.activeSheet = wb.sheets.indexOf(outSheet)
      this.hooks.onActiveSheetChange?.(wb.activeSheet)
      this.rebuild(outSheet, wb, this.view.zoom, this.rendererOpts)
    } else {
      r.rebuildMetrics()
      this.refreshContentSize()
      void this.overlays.build(s, r, this.view)
    }
    this.selectCell(outTop, outLeft)
    if (opts.showPanel) this.showPivotPanel({ sourceSheet: s, outSheet, sourceRange: { ...sel }, outRange: range, fields: fieldOptions, layout: initialLayout })
    return true
  }

  private showPivotPanel(ctx: { sourceSheet: SheetModel; outSheet: SheetModel; sourceRange: MergeRange; outRange: MergeRange; fields: PivotFieldOption[]; layout: PivotTableLayout }): void {
    this.pivotPanel.show({
      fields: ctx.fields,
      filterValues: pivotFilterValues(ctx.sourceSheet, ctx.sourceRange, ctx.fields),
      layout: ctx.layout,
      onChange: (layout) => {
        const table = ctx.outSheet.pivotTables.find((p) => p.range.top === ctx.outRange.top && p.range.left === ctx.outRange.left)
        if (!table) return
        table.layout = clonePivotLayout(layout)
        if (!this.recomputePivot(ctx.outSheet, table)) return
        ctx.outRange.bottom = table.range.bottom
        ctx.outRange.right = table.range.right
        this.renderer?.rebuildMetrics()
        this.refreshContentSize()
        if (this.sheet && this.renderer) void this.overlays.build(this.sheet, this.renderer, this.view)
        this.render()
      },
    })
  }

  /**
   * 透视表的唯一重算入口:读 pivot.source/layout/collapsed → 重建静态结果 → 清旧区写新区 →
   * 刷新 range/buttons/rowGroups。被面板改布局、源数据"活刷新"、折叠/展开三处共用,避免逻辑分叉。
   */
  private recomputePivot(hostSheet: SheetModel, pivot: PivotTableModel): boolean {
    const wb = this.workbook
    if (!wb || !pivot.source || !pivot.layout) return false
    const sourceSheet = wb.sheets[pivot.source.sheetIndex]
    if (!sourceSheet) return false
    const fields = this.pivotFieldOptionsForSheet(sourceSheet, pivot.source.range)
    const built = this.buildPivotRows(sourceSheet, pivot.source.range, fields, pivot.layout, pivot.collapsed ?? [])
    if (!built) return false
    const top = pivot.range.top
    const left = pivot.range.left
    const width = Math.max(1, ...built.rows.map((row) => row.length))
    const newBottom = top + built.rows.length - 1
    const newRight = left + width - 1
    // 清"旧 ∪ 新"区(折叠后新区更小,残留行要清掉),再写新结果
    const clearBottom = Math.max(pivot.range.bottom, newBottom)
    const clearRight = Math.max(pivot.range.right, newRight)
    for (let row = top; row <= clearBottom; row++) {
      for (let col = left; col <= clearRight; col++) hostSheet.cells.delete(cellKey(row, col))
    }
    for (let rr = 0; rr < built.rows.length; rr++) {
      for (let cc = 0; cc < built.rows[rr].length; cc++) setCellValue(hostSheet, top + rr, left + cc, built.rows[rr][cc])
    }
    pivot.range = { top, left, bottom: newBottom, right: newRight }
    pivot.buttons = [
      { row: top, col: left, label: built.rowLabel, kind: 'row' },
      { row: top, col: left + 1, label: built.valueLabel, kind: 'data' },
    ]
    pivot.rowGroups = built.groups.map((g) => ({ row: top + g.rowOffset, key: g.key }))
    return true
  }

  /**
   * 源数据"活刷新":任何模型变更(编辑/粘贴/撤销/重做都经 onModelChange)后,把所有透视表按其源区域重算,
   * 让结果跟着源数据走(WPS 透视表的"活"语义)。重算用 setCellValue 直接改模型(不经命令栈,不再触发
   * onModelChange),pivotRefreshing 兜底防重入。功能开关关闭或无透视表时零开销。
   */
  private refreshPivotsAfterEdit(): void {
    if (this.pivotRefreshing || !this.editCfg.pivotTable) return
    const wb = this.workbook
    if (!wb) return
    let any = false
    for (const sh of wb.sheets) if ((sh.pivotTables?.length ?? 0) > 0) { any = true; break }
    if (!any) return
    this.pivotRefreshing = true
    try {
      for (const hostSheet of wb.sheets) {
        for (const pivot of hostSheet.pivotTables ?? []) {
          if (pivot.source && pivot.layout) this.recomputePivot(hostSheet, pivot)
        }
      }
    } finally {
      this.pivotRefreshing = false
    }
  }

  /** 折叠/展开某个外层行分组(行字段 ≥2 时由折叠按钮触发),重算并重绘。 */
  private togglePivotGroup(tableIdx: number, key: string): void {
    const sheet = this.sheet
    if (!sheet) return
    const pivot = sheet.pivotTables[tableIdx]
    if (!pivot) return
    const collapsed = new Set(pivot.collapsed ?? [])
    if (collapsed.has(key)) collapsed.delete(key)
    else collapsed.add(key)
    pivot.collapsed = [...collapsed]
    if (!this.recomputePivot(sheet, pivot)) return
    this.renderer?.rebuildMetrics()
    this.refreshContentSize()
    if (this.renderer) void this.overlays.build(sheet, this.renderer, this.view)
    this.render()
  }

  /**
   * 重建透视结果为二维单元格数组(框架无关纯计算)。
   * 行字段第 1 个为外层分组,其余为内层明细 → 多行字段时产出可折叠的"大纲"(外层分组行带小计 +
   * 内层缩进明细);单行字段退化为扁平(每个值一行,无折叠)。列字段横向展开,值字段可多个。
   * groups 列出可折叠的分组表头行(相对 top 的偏移 + 外层 key),供渲染折叠按钮 + 命中测试。
   */
  private buildPivotRows(sourceSheet: SheetModel, sel: MergeRange, fields: PivotFieldOption[], layout: PivotTableLayout, collapsed: string[] = []): { rows: Array<Array<string | number>>; rowLabel: string; valueLabel: string; groups: Array<{ rowOffset: number; key: string }> } | null {
    const rowIndexes = layout.rows
    const valueRules = layout.values
    // 空透视表:还没选任何行/值字段 → 占位框(用户在右侧面板勾选字段后会重算填充,对齐 WPS/Excel 空白起步)
    if (!rowIndexes.length && !valueRules.length) {
      return { rows: [['数据透视表']], rowLabel: '数据透视表', valueLabel: '', groups: [] }
    }
    const columns = valueRules.length ? layout.columns : [] // 无值字段时列展开无意义
    const collapsedSet = new Set(collapsed)
    const ruleNumeric = valueRules.map((rule) => fields.find((f) => f.index === rule.field)?.numeric ?? false) // 循环不变量,预存避免每行 find
    const fieldLabel = (index: number) => fields.find((f) => f.index === index)?.label ?? `字段${index + 1}`

    const filteredRows: number[] = []
    for (let row = sel.top + 1; row <= sel.bottom; row++) {
      if (pivotRowKept(sourceSheet, row, layout.filters)) filteredRows.push(row)
    }

    const colKeys = uniqueSorted(filteredRows.map((row) => keyFor(sourceSheet, row, columns)))
    const effectiveColKeys = colKeys.length ? colKeys : ['']
    const valueLabels = valueRules.map((rule) => `${pivotSummaryLabel(rule.summary)}: ${fieldLabel(rule.field)}`)
    const header: Array<string | number> = [rowIndexes.length ? (rowIndexes.map(fieldLabel).join('/') || '行标签') : '总计']
    for (const colKey of effectiveColKeys) {
      for (const label of valueLabels) header.push(colKey ? `${colKey} ${label}` : label)
    }

    // colKey → [每个值规则一个累加器]
    type ColAccs = Map<string, PivotAcc[]>
    const accInto = (ca: ColAccs, colKey: string, row: number): void => {
      let accs = ca.get(colKey)
      if (!accs) ca.set(colKey, (accs = valueRules.map(() => emptyPivotAcc())))
      for (let v = 0; v < valueRules.length; v++) {
        const raw = sourceSheet.cells.get(cellKey(row, valueRules[v].field))?.raw
        const n = pivotNumber(raw)
        addPivotAcc(accs[v], ruleNumeric[v] && n != null ? n : null, raw)
      }
    }
    const lineOf = (label: string, ca: ColAccs): Array<string | number> => {
      const line: Array<string | number> = [label]
      effectiveColKeys.forEach((colKey) => {
        const accs = ca.get(colKey)
        for (let v = 0; v < valueRules.length; v++) line.push(accs ? resolvePivotAcc(accs[v], valueRules[v].summary) : 0)
      })
      return line
    }

    // 只有值字段(无行字段)→ 整体一行总计
    if (!rowIndexes.length) {
      const ca: ColAccs = new Map()
      for (const row of filteredRows) accInto(ca, columns.length ? keyFor(sourceSheet, row, columns) : '', row)
      return { rows: [header, lineOf('总计', ca)], rowLabel: header[0] as string, valueLabel: valueLabels.join('/'), groups: [] }
    }

    const outerField = rowIndexes[0]
    const innerFields = rowIndexes.slice(1)
    const hasHierarchy = innerFields.length > 0
    interface OuterGroup { subtotal: ColAccs; inner: Map<string, ColAccs> }
    const outerGroups = new Map<string, OuterGroup>()
    for (const row of filteredRows) {
      const outerKey = keyFor(sourceSheet, row, [outerField])
      const colKey = columns.length ? keyFor(sourceSheet, row, columns) : ''
      let g = outerGroups.get(outerKey)
      if (!g) outerGroups.set(outerKey, (g = { subtotal: new Map(), inner: new Map() }))
      accInto(g.subtotal, colKey, row)
      if (hasHierarchy) {
        const innerKey = keyFor(sourceSheet, row, innerFields)
        let ic = g.inner.get(innerKey)
        if (!ic) g.inner.set(innerKey, (ic = new Map()))
        accInto(ic, colKey, row)
      }
    }
    if (!outerGroups.size) return null

    const grand: PivotAcc[][] = effectiveColKeys.map(() => valueRules.map(() => emptyPivotAcc()))
    const rows: Array<Array<string | number>> = [header]
    const groups: Array<{ rowOffset: number; key: string }> = []
    for (const outerKey of uniqueSorted([...outerGroups.keys()])) {
      const g = outerGroups.get(outerKey)!
      effectiveColKeys.forEach((colKey, ci) => {
        const accs = g.subtotal.get(colKey)
        if (accs) for (let v = 0; v < valueRules.length; v++) mergePivotAcc(grand[ci][v], accs[v])
      })
      if (hasHierarchy) {
        groups.push({ rowOffset: rows.length, key: outerKey })
        rows.push(lineOf(`  ${outerKey}`, g.subtotal)) // 分组表头(=小计);前导空格给折叠按钮让位
        if (!collapsedSet.has(outerKey)) {
          for (const innerKey of uniqueSorted([...g.inner.keys()])) rows.push(lineOf(`　　${innerKey}`, g.inner.get(innerKey)!))
        }
      } else {
        rows.push(lineOf(outerKey, g.subtotal)) // 单行字段:扁平,每个值一行
      }
    }
    const grandLine: Array<string | number> = ['总计']
    for (let ci = 0; ci < effectiveColKeys.length; ci++) for (let v = 0; v < valueRules.length; v++) grandLine.push(resolvePivotAcc(grand[ci][v], valueRules[v].summary))
    rows.push(grandLine)
    return { rows, rowLabel: header[0] as string, valueLabel: valueLabels.join('/'), groups }
  }

  private failPivot(message: string): false {
    console.warn(`[ooxml-preview] ${message}`)
    if (typeof window !== 'undefined') window.alert(message)
    return false
  }

  private pivotFieldOptions(sel: MergeRange): PivotFieldOption[] {
    const s = this.sheet
    if (!s) return []
    return this.pivotFieldOptionsForSheet(s, sel)
  }

  private pivotFieldOptionsForSheet(s: SheetModel, sel: MergeRange): PivotFieldOption[] {
    if (sel.bottom <= sel.top) return []
    const out: PivotFieldOption[] = []
    for (let col = sel.left; col <= sel.right; col++) {
      let numeric = 0
      for (let row = sel.top + 1; row <= sel.bottom; row++) {
        if (pivotNumber(s.cells.get(cellKey(row, col))?.raw) != null) numeric++
      }
      out.push({ index: col, label: this.pivotCellText(s, sel.top, col) || `字段${col - sel.left + 1}`, numeric: numeric > 0 })
    }
    return out
  }

  private pivotCellText(sheet: SheetModel, row: number, col: number): string {
    if (sheet === this.sheet && this.renderer) return this.renderer.cellText(row, col).trim()
    const raw = sheet.cells.get(cellKey(row, col))?.raw
    return raw == null ? '' : String(raw).trim()
  }

  /**
   * 按某列对自动筛选数据区排序(只读视图重排:移动行内容 + 行高,不改文件)。
   * 合并区与数据区相交则拒绝(与 Excel 一致)。排序前会清除该表筛选,避免行索引快照错位。
   */
  sortColumn(col: number, dir: 'asc' | 'desc'): void {
    const r = this.renderer
    const s = this.sheet
    if (!r || !s?.autoFilterRange) return
    const af = s.autoFilterRange
    const top = af.top + 1
    const bottom = this.filterDataBottom()
    const left = af.left
    const right = af.right
    if (bottom <= top) return // 不足两行数据,无需排序
    if (col < left || col > right) return
    // 合并区相交 → 拒绝(Excel 行为)
    for (const m of s.merges) {
      if (!(m.bottom < top || m.top > bottom || m.right < left || m.left > right)) {
        console.warn('[ooxml-preview] 排序区域含合并单元格,已跳过(与 Excel 一致)')
        return
      }
    }
    // 先清筛选(恢复隐藏行 + 清状态),避免按行索引的快照在重排后错位
    this.resetFilter(s)

    // 收集每行 [left..right] 的单元格 + 行信息 + 排序键
    type Entry = { key: CellModel['raw']; cells: CellModel[]; rowInfo: RowInfo | undefined }
    const entries: Entry[] = []
    for (let row = top; row <= bottom; row++) {
      const cells: CellModel[] = []
      for (let c = left; c <= right; c++) {
        const cell = s.cells.get(cellKey(row, c))
        if (cell) cells.push(cell)
      }
      const keyCell = s.cells.get(cellKey(row, col))
      entries.push({ key: keyCell ? keyCell.raw : null, cells, rowInfo: s.rows.get(row) })
    }
    // 稳定排序:空值恒排末尾(不随方向翻转),其余按类型/数值/文本比较
    const sign = dir === 'asc' ? 1 : -1
    const order = entries.map((_, i) => i)
    order.sort((a, b) => {
      const ka = entries[a].key
      const kb = entries[b].key
      const ba = isBlankValue(ka)
      const bb = isBlankValue(kb)
      if (ba && bb) return a - b
      if (ba) return 1
      if (bb) return -1
      const c = compareCellValues(ka, kb)
      return c !== 0 ? sign * c : a - b
    })
    // 清空旧区段,再按新顺序回填
    for (let row = top; row <= bottom; row++) {
      for (let c = left; c <= right; c++) s.cells.delete(cellKey(row, c))
      s.rows.delete(row)
    }
    for (let i = 0; i < order.length; i++) {
      const targetRow = top + i
      const e = entries[order[i]]
      for (const cell of e.cells) {
        cell.row = targetRow
        s.cells.set(cellKey(targetRow, cell.col), cell)
      }
      if (e.rowInfo) s.rows.set(targetRow, e.rowInfo)
    }
    this.sortCol = col
    this.sortDir = dir
    r.rebuildMetrics()
    this.refreshContentSize()
    this.clearSelection()
    this.hooks.onFilterChange() // 浮层/工具栏据此重算排序指示
    this.render()
  }

  // ====================== 编辑配置(E0:闸门) ======================

  /** 设置编辑配置(默认只读;壳在挂载 + props 变化时调) */
  setEditConfig(cfg: EditConfig): void {
    this.editCfg = cfg ?? {}
    this.edit.setPasteBehavior(this.editCfg.pasteBehavior ?? null) // 粘贴行为默认(缺项回落默认)
    this.edit.refreshEngine() // recalc/formulaEngine 可能变了 → 重置引擎并按需点火
  }

  /** 该格当前是否可编辑(综合 editable + editableTargets 白名单 + readOnlyRanges + cellReadOnly) */
  isCellEditable(row: number, col: number): boolean {
    return this.sheet ? resolveEditable(this.sheet, row, col, this.editCfg) : false
  }

  /**
   * **运行时**改可编辑白名单(2026-06-08 新增) —— 不动 `:editableTargets` prop,
   * 直接覆盖 `editCfg.editableTargets`. 立即重绘以反映只读光标变化.
   * 传 `undefined` = 关闭白名单(默认全可编辑);`[]` = 全只读;单值或数组 = 白名单.
   */
  setEditableTargets(targets: EditConfig['editableTargets']): void {
    this.editCfg = { ...this.editCfg, editableTargets: targets }
    this.render() // editable 状态变了, 重绘以触发外部 cell-style 钩子重算 (可视化只读/可编辑)
  }

  /** 当前生效的可编辑白名单(运行时 setEditableTargets 或初始 prop). 用 `undefined` 表示未启用白名单 */
  getEditableTargets(): EditConfig['editableTargets'] {
    return this.editCfg.editableTargets
  }

  // ---- 命令式编辑 API(委托 EditController;E1) ----
  editCell(row: number, col: number, value: CellValue): boolean {
    return this.edit.editCell(row, col, value)
  }
  editRange(range: MergeRange, values: CellValue[][]): boolean {
    return this.edit.editRange(range, values)
  }
  clearRange(range: MergeRange): boolean {
    return this.edit.clearRange(range)
  }
  /** 给区域套样式覆盖(E5;粗体/对齐/填充等);editable 时走命令栈(可撤销 + 发 cell-change + 记脏) */
  setStyle(range: MergeRange, patch: CellStyleOverride): boolean {
    return this.edit.setStyle(range, patch)
  }

  // ---- 背景色 / 字体色 回显 + 修改(WPS 风格工具栏用) ----
  /** 活动格当前背景填充色(#RRGGBB);无填充/非纯色 → 默认白 #FFFFFF。 */
  getActiveFillColor(): string {
    const a = this.selActive
    const snap = a ? this.edit.getCellSnapshot(a.row, a.col) : null
    const fill = snap?.style?.fill
    return (fill && fill.type === 'solid' && toHex6(fill.fgColor)) || '#FFFFFF'
  }
  /** 活动格当前字体色(#RRGGBB);缺省黑 #000000。 */
  getActiveFontColor(): string {
    const a = this.selActive
    const snap = a ? this.edit.getCellSnapshot(a.row, a.col) : null
    return toHex6(snap?.style?.font?.color) || '#000000'
  }
  /** 给当前选区设背景填充色(null = 清除填充);editable 时入命令栈。 */
  setSelectionFill(color: string | null): boolean {
    const sel = this.getSelection()
    if (!sel) return false
    return this.setStyle(sel, { fill: color ? { type: 'solid', fgColor: color } : { type: 'none' } })
  }
  /** 给当前选区设字体色;editable 时入命令栈。 */
  setSelectionFontColor(color: string): boolean {
    const sel = this.getSelection()
    if (!sel) return false
    return this.setStyle(sel, { font: { color } })
  }

  // ---- 自动换行(WPS 风格 toggle) ----
  /** 当前选区里 wrapText 的整体态:'all' 全开 / 'none' 全关 / 'mixed' 混合。空选区→'none'。 */
  getSelectionWrapState(): 'all' | 'none' | 'mixed' {
    const sel = this.getSelection()
    if (!sel || !this.sheet) return 'none'
    let yes = 0
    let no = 0
    for (let r = sel.top; r <= sel.bottom; r++) {
      for (let c = sel.left; c <= sel.right; c++) {
        const snap = this.edit.getCellSnapshot(r, c)
        if (snap?.style?.wrapText) yes++
        else no++
        if (yes && no) return 'mixed'
      }
    }
    return yes ? 'all' : 'none'
  }
  /** 切换当前选区的"自动换行"(WPS 风格):全 wrap → 全关;否则 → 全开。
   *  失效行高缓存让 autofit 按新 wrap 重撑(只扩不缩);editable 时入命令栈(单次撤销 style)。
   *  注:undo 回滚 style 但不回滚行高,与现有 setStyle/autofit "只扩不缩"语义一致。 */
  toggleWrapTextOnSelection(): boolean {
    if (!this.editCfg.editable) return false
    const sel = this.getSelection()
    if (!sel || !this.sheet) return false
    const target = this.getSelectionWrapState() !== 'all'
    invalidateAutofit(this.sheet) // 必须先失效,setStyle 会同步触发 render→autofit
    return this.setStyle(sel, { wrapText: target })
  }
  /** 合并区域(G1;清空被覆盖格,只留左上锚点);editable 时入命令栈 */
  mergeCells(range: MergeRange): boolean {
    return this.edit.mergeCells(range)
  }
  /** 拆分区域内的合并(G1);editable 时入命令栈 */
  unmergeCells(range: MergeRange): boolean {
    return this.edit.unmergeCells(range)
  }

  // ---- 图片编辑(E6;浮动/嵌入 增删移改) ----
  /** 读当前表全部图片锚点(克隆)。 */
  getImages(): ImageAnchor[] {
    return this.edit.getImages()
  }
  /** 加一张图(无 src 但有 bytes+mime 时自动生成 blob url);返回插入索引(失败 -1)。 */
  addImage(anchor: ImageAnchor): number {
    let a = anchor
    if (!a.src && a.bytes && a.mime) {
      a = { ...a, src: URL.createObjectURL(new Blob([a.bytes as BlobPart], { type: a.mime })) }
    }
    return this.edit.addImage(a)
  }
  /** 删一张图。 */
  removeImage(index: number): boolean {
    return this.edit.removeImage(index)
  }
  /** 移动图片(屏幕像素增量);editable 时入命令栈 + 发 image-change。 */
  moveImage(index: number, dxPx: number, dyPx: number): boolean {
    return this.editImageRect(index, (rect) => ({ ...rect, left: rect.left + dxPx, top: rect.top + dyPx }))
  }
  /** 缩放图片(目标屏幕像素宽高);editable 时入命令栈 + 发 image-change。 */
  resizeImage(index: number, widthPx: number, heightPx: number): boolean {
    return this.editImageRect(index, (rect) => ({ ...rect, width: Math.max(8, widthPx), height: Math.max(8, heightPx) }))
  }

  // ---- WPS 单元格内嵌图(DISPIMG)⇄ 浮动图互转(第二期) ----
  /** 改 WPS 单元格内嵌图贴合方式(fill/contain/cover);即时重绘。 */
  setCellImageFit(fit: CellImageFit): void {
    this.rendererOpts = { ...this.rendererOpts, cellImageFit: fit }
    if (this.renderer?.setCellImageFit(fit)) this.render()
  }
  /** 读 WPS 单元格内嵌图登记表(克隆,id→{id,src,mime});无则空数组。 */
  getCellImages(): { id: string; src: string; mime?: string }[] {
    const reg = this.workbook?.cellImages
    if (!reg) return []
    return Array.from(reg.values(), (ci) => ({ id: ci.id, src: ci.src, mime: ci.mime }))
  }
  /** 某格是否内嵌图(DISPIMG)→ 返 {id,src,mime},否则 null(供点击放大判定)。 */
  getCellImageAt(row: number, col: number): { id: string; src: string; mime?: string } | null {
    const id = this.sheet?.cells.get(cellKey(row, col))?.dispImgId
    const ci = id ? this.workbook?.cellImages?.get(id) : undefined
    return ci ? { id: ci.id, src: ci.src, mime: ci.mime } : null
  }
  /** 开/关图片点击放大灯箱(默认开;只读模式单击图、编辑模式右键菜单触发)。 */
  setLightboxEnabled(b: boolean): void {
    this.lightboxEnabled = b
  }
  /** 打开图片放大灯箱(命令式;src = blob/data/http url)。 */
  openImageLightbox(src: string, fileName?: string, mime?: string): void {
    this.lightbox.show({ src, fileName, mime })
  }
  /** 一张浮动图视觉中心落在哪个单元格(用几何反推;无渲染器时回落锚点 from 格)。 */
  imageCellOf(index: number): { row: number; col: number } | null {
    const img = this.sheet?.images[index]
    if (!img) return null
    if (!this.renderer) return { row: img.from.row, col: img.from.col }
    const r = anchorRect(this.renderer.metrics, img)
    return {
      row: this.renderer.metrics.rowAt(r.top + r.height / 2),
      col: this.renderer.metrics.colAt(r.left + r.width / 2),
    }
  }
  /** 浮动图 → 单元格内嵌图(显式指定目标格);失败返 false。 */
  convertImageToCell(imageIndex: number, row: number, col: number): boolean {
    return this.edit.convertImageToCell(imageIndex, row, col)
  }
  /** 选区批量:把"中心格落在 range 内"的所有浮动图就近嵌入,聚合成单次 undo;返回成功嵌入张数。 */
  convertImagesInRangeToCell(range: MergeRange): number {
    const sheet = this.sheet
    if (!sheet) return 0
    const targets: { imageIndex: number; row: number; col: number }[] = []
    for (let i = 0; i < sheet.images.length; i++) {
      const at = this.imageCellOf(i)
      if (!at) continue
      if (at.row >= range.top && at.row <= range.bottom && at.col >= range.left && at.col <= range.right) {
        targets.push({ imageIndex: i, row: at.row, col: at.col })
      }
    }
    if (!targets.length) return 0
    return this.edit.convertImagesToCells(targets)
  }
  // applyTemplate(P3 旧版 placeholder + anchor 模型,2026-06-08 已删除)
  // 新语义"模板=样式捐赠者"由 core/template/style-overlay.ts 提供,壳层 (Vue/React) 用
  // applyStyleTemplate(dataWb, templateWb) 合成后再 loadModel 喂进来,不再走 controller 内 in-place 改造.
  /** 选区批量:把 range 内所有 DISPIMG 格拎成浮动图;聚合成单次 undo;返回成功转换张数。 */
  convertCellImagesInRangeToFloat(range: MergeRange, size?: { width: number; height: number }): number {
    const sheet = this.sheet
    if (!sheet) return 0
    const cells: { row: number; col: number; size?: { width: number; height: number } }[] = []
    for (let r = range.top; r <= range.bottom; r++) {
      for (let c = range.left; c <= range.right; c++) {
        const cell = sheet.cells.get(cellKey(r, c))
        if (cell?.dispImgId) cells.push({ row: r, col: c, size })
      }
    }
    if (!cells.length) return 0
    return this.edit.convertCellImagesToFloats(cells)
  }
  /** 浮动图 → 内嵌图(**就近**:图在哪格就嵌哪格,目标由几何反推);失败返 false。 */
  convertImageToCellAuto(imageIndex: number): boolean {
    const at = this.imageCellOf(imageIndex)
    return at ? this.edit.convertImageToCell(imageIndex, at.row, at.col) : false
  }
  /**
   * 批量把浮动图按"所在单元格"就近嵌入(整表 / 整列)。`col` 给定则只嵌中心落在该列的图。
   * 一次进撤销栈(单次 Ctrl+Z 全撤)。返回成功嵌入的张数。
   */
  convertAllImagesToCells(col?: number): number {
    const imgs = this.sheet?.images
    if (!imgs || !imgs.length) return 0
    const targets: { imageIndex: number; row: number; col: number }[] = []
    for (let i = 0; i < imgs.length; i++) {
      const at = this.imageCellOf(i)
      if (!at) continue
      if (col != null && at.col !== col) continue
      targets.push({ imageIndex: i, row: at.row, col: at.col })
    }
    return this.edit.convertImagesToCells(targets)
  }
  /** 单元格内嵌图 → 浮动图(把 row,col 的 DISPIMG 拎成浮动图);非内嵌图格返 false。 */
  convertCellImageToFloat(row: number, col: number, size?: { width: number; height: number }): boolean {
    return this.edit.convertCellImageToFloat(row, col, size)
  }
  // ---- 行列结构编辑(E7;增删行列) ----
  /** 在 at 处插入 count 行(原 at 行及之后下移)。 */
  insertRows(at: number, count = 1): boolean {
    return this.edit.insertRows(at, count)
  }
  /** 删除 [at, at+count) 行(与合并相交则警告,相交合并被移除)。 */
  deleteRows(at: number, count = 1): boolean {
    if (this.sheet && deleteIntersectsMerge(this.sheet, 'delete-rows', at, count))
      console.warn('[ooxml-preview] 删除行与合并单元格相交,相交的合并将被移除')
    return this.edit.deleteRows(at, count)
  }
  /** 在 at 处插入 count 列。 */
  insertCols(at: number, count = 1): boolean {
    return this.edit.insertCols(at, count)
  }
  /** 删除 [at, at+count) 列(与合并相交则警告)。 */
  deleteCols(at: number, count = 1): boolean {
    if (this.sheet && deleteIntersectsMerge(this.sheet, 'delete-cols', at, count))
      console.warn('[ooxml-preview] 删除列与合并单元格相交,相交的合并将被移除')
    return this.edit.deleteCols(at, count)
  }

  /** 移动/缩放公共路径:算当前矩形 → 变换 → setImageRect → 重定位 → 补登命令。 */
  private editImageRect(
    index: number,
    fn: (rect: { left: number; top: number; width: number; height: number }) => { left: number; top: number; width: number; height: number },
  ): boolean {
    if (!this.editCfg.editable || !this.sheet || !this.renderer) return false
    const img = this.sheet.images[index]
    if (!img) return false
    this.edit.ensureBaseline()
    const before = cloneImageAnchor(img)
    setImageRect(this.sheet, index, fn(anchorRect(this.renderer.metrics, img)), this.view.zoom)
    this.render()
    this.edit.recordImageEdit(index, before, cloneImageAnchor(this.sheet.images[index]))
    return true
  }
  undo(): void {
    this.edit.undo()
  }
  redo(): void {
    this.edit.redo()
  }
  canUndo(): boolean {
    return this.edit.canUndo()
  }
  canRedo(): boolean {
    return this.edit.canRedo()
  }
  getEditingCell(): { row: number; col: number } | null {
    return this.edit.getEditingCell()
  }
  getCellSnapshot(row: number, col: number): CellSnapshot | null {
    return this.edit.getCellSnapshot(row, col)
  }
  /** 单元格"全息体检":snapshot + 合并区 + 浮动图覆盖 + WPS 内嵌图 + 数据验证 + 条件格式命中 + 链接/批注。
   *  无 workbook / sheet / 越界返 null。详见 [src/core/model/inspect.ts](src/core/model/inspect.ts)。 */
  inspectCell(row: number, col: number): CellInspection | null {
    const sheet = this.sheet
    const wb = this.workbook
    if (!sheet || !wb) return null
    if (row < 0 || col < 0) return null
    return inspectCell(sheet, wb, row, col, wb.date1904)
  }

  // ---- 编辑器宿主(E2) ----
  /** 壳注入合并后的编辑器解析器(plugin.editor + prop.editor) */
  setEditorResolver(fn?: EditorResolver): void {
    this.editorResolver = fn
  }

  /**
   * 进入编辑:解析编辑器工厂(无自定义则用内置文本编辑器)→ 挂载。只读则不进入。
   * initialText 为打字进入时的起始字符。返回是否进入。
   */
  beginEdit(row: number, col: number, initialText?: string): boolean {
    if (!this.sheet || !this.workbook) return false
    // 合并单元格 → 编辑落到锚点(左上格),编辑框盖住整片合并区(像 WPS/Excel)
    const merge = this.renderer?.mergeAt(row, col) ?? null
    const ar = merge ? merge.top : row
    const ac = merge ? merge.left : col
    if (!this.isCellEditable(ar, ac)) return false
    const cell = this.sheet.cells.get(cellKey(ar, ac)) ?? null
    const factory = this.editorResolver?.(cell, { row: ar, col: ac }) ?? defaultCellEditor // E3:内置兜底
    const rect = merge ? this.rectOfRange(merge) : this.rectOf(ar, ac)
    const snapshot = this.edit.getCellSnapshot(ar, ac)
    if (!rect || !snapshot) return false
    const ctx: CellEditorContext = {
      snapshot,
      rect,
      sheet: this.sheet,
      workbook: this.workbook,
      permission: 'editable',
      initialText,
      commit: (v, move) => this.commitEdit(v, move),
      cancel: () => this.cancelEdit(),
    }
    const rectOverride = merge ? () => this.rectOfRange(merge) : undefined
    if (!this.editorHost.mount(ar, ac, factory, ctx, rectOverride)) return false
    this.edit.setEditing({ row: ar, col: ac })
    this.hooks.onEditEvent('edit-start', { cell: { row: ar, col: ac }, snapshot })
    return true
  }

  /**
   * 提交当前编辑(取值 → 命令栈 → 卸编辑器)。value 可为裸值或 {value,style}(样式 E5 起生效)。
   * move 指示提交后活动格移动(Enter→down / Tab→right)。
   */
  commitEdit(value: EditorCommitValue, move?: 'down' | 'right'): void {
    const editing = this.edit.getEditingCell()
    if (!editing) return
    const wrapped = value !== null && typeof value === 'object' && !(value instanceof Date) && 'value' in value
    const val: CellValue = wrapped ? (value as { value: CellValue }).value : (value as CellValue)
    const style = wrapped ? (value as { style?: CellStyleOverride }).style : undefined
    this.edit.editCell(editing.row, editing.col, val)
    // E5:编辑器可返 { value, style } → 顺带套自定义编辑样式(要求 2 端到端)
    if (style) this.edit.setStyle({ top: editing.row, left: editing.col, bottom: editing.row, right: editing.col }, style)
    this.hooks.onEditEvent('edit-commit', { cell: editing, value: val })
    this.editorHost.unmount()
    this.edit.setEditing(null)
    // 提交后导航 + 让滚动容器重新拿焦点(以便继续键盘操作)
    const r = this.renderer
    if (move && r) {
      const nr = move === 'down' ? Math.min(editing.row + 1, r.metrics.rows - 1) : editing.row
      const nc = move === 'right' ? Math.min(editing.col + 1, r.metrics.cols - 1) : editing.col
      this.selectCell(nr, nc)
    }
    this.els.scroller.focus()
  }

  /** 取消当前编辑(不改模型) */
  cancelEdit(): void {
    if (!this.editorHost.isActive()) return
    this.editorHost.unmount()
    this.edit.setEditing(null)
  }

  /** 当前是否有活动编辑器 */
  isEditing(): boolean {
    return this.editorHost.isActive()
  }

  // ====================== 导出 / 打印(委托 WorkbookExporter) ======================

  exportImage(opts?: ImageExportOptions): Promise<Blob> {
    return this.exporter.exportImage(opts)
  }
  downloadImage(opts?: ImageExportOptions): Promise<void> {
    return this.exporter.downloadImage(opts)
  }
  exportPdf(opts?: PdfExportOptions): Promise<Blob> {
    return this.exporter.exportPdf(opts)
  }
  downloadPdf(opts?: PdfExportOptions): Promise<void> {
    return this.exporter.downloadPdf(opts)
  }
  print(opts?: PrintOptions): Promise<void> {
    return this.exporter.print(opts)
  }
  // ---- 数据导出(E8;委托 WorkbookExporter,一份数据层 → xlsx/json/csv) ----
  exportXlsx(opts?: XlsxExportOptions): Promise<Blob> {
    return this.exporter.exportXlsx(opts)
  }
  downloadXlsx(opts?: XlsxExportOptions): Promise<void> {
    return this.exporter.downloadXlsx(opts)
  }
  exportJson(opts?: SheetToJSONOptions): string {
    return this.exporter.exportJson(opts)
  }
  downloadJson(opts?: SheetToJSONOptions): void {
    this.exporter.downloadJson(opts)
  }
  exportCsv(opts?: { target?: number; format?: boolean }): string {
    return this.exporter.exportCsv(opts)
  }
  downloadCsv(opts?: { target?: number; format?: boolean }): void {
    this.exporter.downloadCsv(opts)
  }
}

// ---- 命中区域 ----
type Hit =
  | { region: 'cell'; row: number; col: number }
  | { region: 'row'; row: number }
  | { region: 'col'; col: number }
  | { region: 'corner' }
  | { region: 'none' }

function cellRange(c: Cell): MergeRange {
  return { top: c.row, left: c.col, bottom: c.row, right: c.col }
}
function colLabel(col: number): string {
  let n = col + 1
  let s = ''
  while (n > 0) {
    const r = (n - 1) % 26
    s = String.fromCharCode(65 + r) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

function parseCellRef(ref: string): Cell | null {
  const m = /^\s*([A-Za-z]+)([1-9]\d*)\s*$/.exec(ref)
  if (!m) return null
  let col = 0
  for (const ch of m[1].toUpperCase()) col = col * 26 + ch.charCodeAt(0) - 64
  const row = Number(m[2]) - 1
  col -= 1
  return row >= 0 && row < MAX_GRID_ROWS && col >= 0 && col < MAX_GRID_COLS ? { row, col } : null
}

function pivotNumber(raw: CellModel['raw'] | undefined): number | null {
  if (typeof raw === 'number' && Number.isFinite(raw)) return raw
  if (typeof raw === 'string') {
    const n = Number(raw.replace(/,/g, '').trim())
    if (Number.isFinite(n)) return n
  }
  return null
}

function keyFor(sheet: SheetModel, row: number, cols: number[]): string {
  if (!cols.length) return ''
  return cols.map((col) => {
    const raw = sheet.cells.get(cellKey(row, col))?.raw
    return raw == null || String(raw).trim() === '' ? '(空白)' : String(raw)
  }).join(' / ')
}

function uniqueSorted(values: string[]): string[] {
  return [...new Set(values)].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
}

function pivotFilterValues(sheet: SheetModel, sel: MergeRange, fields: PivotFieldOption[]): Record<number, string[]> {
  const out: Record<number, string[]> = {}
  for (const field of fields) {
    const values: string[] = []
    for (let row = sel.top + 1; row <= sel.bottom; row++) {
      const raw = sheet.cells.get(cellKey(row, field.index))?.raw
      if (raw != null && String(raw).trim() !== '') values.push(String(raw))
    }
    out[field.index] = uniqueSorted(values)
  }
  return out
}

function normalizePivotLayout(layout: Partial<PivotTableLayout> | undefined, rowField: number, valueField: number): PivotTableLayout {
  return {
    filters: layout?.filters?.map((rule) => ({ ...rule, values: rule.values?.slice() })) ?? [],
    columns: layout?.columns?.slice() ?? [],
    rows: layout?.rows?.length ? layout.rows.slice() : [rowField],
    values: layout?.values?.length ? layout.values.map((rule) => ({ ...rule })) : [{ field: valueField, summary: 'sum' }],
  }
}

function clonePivotLayout(layout: PivotTableLayout): PivotTableLayout {
  return {
    filters: layout.filters.map((rule) => ({ ...rule, values: rule.values?.slice() })),
    columns: layout.columns.slice(),
    rows: layout.rows.slice(),
    values: layout.values.map((rule) => ({ ...rule })),
  }
}

/** 一行是否通过所有筛选规则(WPS 语义:all 不过滤 / non-empty 去空 / equals 单值 / include 多选包含)。 */
function pivotRowKept(sheet: SheetModel, row: number, filters: PivotFilterRule[]): boolean {
  for (const rule of filters) {
    const raw = sheet.cells.get(cellKey(row, rule.field))?.raw
    const value = raw == null || String(raw).trim() === '' ? '' : String(raw)
    if (rule.mode === 'non-empty' && !value) return false
    if (rule.mode === 'equals' && value !== (rule.value ?? '')) return false
    if (rule.mode === 'include' && rule.values && rule.values.length && !rule.values.includes(value)) return false
  }
  return true
}

function createPivotSheet(wb: WorkbookModel): SheetModel {
  const base = 'PivotTable'
  const used = new Set(wb.sheets.map((s) => s.name))
  let name = base
  let n = 1
  while (used.has(name)) name = `${base}${++n}`
  const sheet: SheetModel = {
    name,
    index: wb.sheets.length,
    state: 'visible',
    dimension: { rows: 0, cols: 0 },
    cells: new Map(),
    styles: [defaultPivotStyle()],
    merges: [],
    columns: new Map(),
    rows: new Map(),
    defaultColWidth: 80,
    defaultRowHeight: 24,
    freeze: { frozenRows: 0, frozenCols: 0 },
    conditional: [],
    dataValidations: [],
    images: [],
    charts: [],
    shapes: [],
    sparklines: [],
    pivotTables: [],
    showGridLines: true,
  }
  wb.sheets.push(sheet)
  return sheet
}

function defaultPivotStyle(): CellStyle {
  return {
    font: { name: 'Calibri', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#000000' },
    fill: { type: 'none' },
    borders: {},
    hAlign: 'general',
    vAlign: 'bottom',
    wrapText: false,
    shrinkToFit: false,
    textRotation: 0,
    indent: 0,
    numFmt: 'General',
  }
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** 剪贴板文本 → 2D 网格(行用换行、列用制表符;去掉末尾换行)。空文本 → []。 */
function parseClipboardGrid(text: string): string[][] {
  if (!text) return []
  let t = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
  if (t.endsWith('\n')) t = t.slice(0, -1)
  if (t === '') return []
  return t.split('\n').map((line) => line.split('\t'))
}

/** 排序用:空值判定(null / 空串 / undefined) */
export function isBlankValue(v: CellModel['raw']): boolean {
  return v === null || v === undefined || v === ''
}

/**
 * 排序用:非空值比较(升序基准,返回 -1/0/1)。
 * 类型序:数字/日期 < 文本 < 布尔;同类型内数值比大小、文本按自然顺序(中文/数字混排)。
 */
export function compareCellValues(a: CellModel['raw'], b: CellModel['raw']): number {
  const ra = typeRank(a)
  const rb = typeRank(b)
  if (ra !== rb) return ra - rb
  if (ra === 0) {
    const na = toNumber(a)
    const nb = toNumber(b)
    return na < nb ? -1 : na > nb ? 1 : 0
  }
  if (ra === 2) return (a ? 1 : 0) - (b ? 1 : 0)
  return String(a).localeCompare(String(b), undefined, { numeric: true })
}

type PivotAcc = { sum: number; count: number; max: number | null; min: number | null }

function emptyPivotAcc(): PivotAcc {
  return { sum: 0, count: 0, max: null, min: null }
}

function addPivotAcc(acc: PivotAcc, numeric: number | null, raw: CellModel['raw'] | undefined): void {
  if (numeric != null) {
    acc.sum += numeric
    acc.count++
    acc.max = acc.max == null ? numeric : Math.max(acc.max, numeric)
    acc.min = acc.min == null ? numeric : Math.min(acc.min, numeric)
  } else if (raw != null && String(raw).trim()) {
    acc.count++
  }
}

function mergePivotAcc(target: PivotAcc, source: PivotAcc): void {
  target.sum += source.sum
  target.count += source.count
  if (source.max != null) target.max = target.max == null ? source.max : Math.max(target.max, source.max)
  if (source.min != null) target.min = target.min == null ? source.min : Math.min(target.min, source.min)
}

function resolvePivotAcc(acc: PivotAcc, summary: PivotSummary): number {
  if (summary === 'count') return acc.count
  if (summary === 'avg') return acc.count ? acc.sum / acc.count : 0
  if (summary === 'max') return acc.max ?? 0
  if (summary === 'min') return acc.min ?? 0
  return acc.sum
}

function pivotSummaryLabel(summary: PivotSummary): string {
  const labels: Record<PivotSummary, string> = { sum: '求和项', count: '计数项', avg: '平均值', max: '最大值', min: '最小值' }
  return labels[summary]
}

function typeRank(v: CellModel['raw']): number {
  if (typeof v === 'number' || v instanceof Date) return 0
  if (typeof v === 'boolean') return 2
  return 1
}
function toNumber(v: CellModel['raw']): number {
  return v instanceof Date ? v.getTime() : typeof v === 'number' ? v : Number(v)
}

/** Ctrl+方向: 跳到数据块边界(Excel 行为) */
function jumpEdge(
  r: CanvasRenderer,
  row: number,
  col: number,
  dr: number,
  dc: number,
): { row: number; col: number } {
  const maxRow = r.metrics.rows - 1
  const maxCol = r.metrics.cols - 1
  const inB = (rr: number, cc: number) => rr >= 0 && rr <= maxRow && cc >= 0 && cc <= maxCol
  const filled = (rr: number, cc: number) => r.cellText(rr, cc) !== ''
  let nr = row + dr
  let nc = col + dc
  if (!inB(nr, nc)) return { row, col }
  if (filled(row, col) && filled(nr, nc)) {
    // 沿填充块走到块尾
    while (inB(nr + dr, nc + dc) && filled(nr + dr, nc + dc)) {
      nr += dr
      nc += dc
    }
  } else {
    // 跳过空白到下一个填充(或边界)
    while (inB(nr, nc) && !filled(nr, nc)) {
      if (!inB(nr + dr, nc + dc)) break
      nr += dr
      nc += dc
    }
  }
  return { row: nr, col: nc }
}
