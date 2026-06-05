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
import type { CellModel, CellStyleOverride, ColumnInfo, ImageAnchor, MergeRange, RowInfo, SheetModel, WorkbookModel } from '../model/types'
import { cellKey } from '../model/types'
import { setImageRect, cloneImageAnchor } from '../model/mutations'
import { deleteIntersectsMerge } from '../model/structure'
import { anchorRect } from '../overlay/anchor'
import type { EditConfig } from '../edit/types'
import { resolveEditable } from '../edit/permissions'
import { EditController, type EditControllerHost, type EditEventName } from '../edit/edit-controller'
import { defaultFormulaEngineFactory } from '../formula/hyperformula-adapter'
import type { CellValue, SheetToJSONOptions } from '../model/data-access'
import type { CellSnapshot } from '../model/snapshot'
import { CellEditorHost } from '../edit/editor-host'
import type { CellEditorContext, EditorCommitValue, EditorResolver } from '../edit/editor-context'
import { defaultCellEditor } from '../edit/default-editor'
import { CanvasRenderer, type RendererOptions, type ViewState } from '../render/canvas-renderer'
import { OverlayManager, type OverlayQuads } from './overlay-manager'
import { WorkbookExporter, type ExporterHost } from '../export/exporter'
import type { ImageExportOptions, PdfExportOptions, PrintOptions } from '../export/types'
import type { XlsxExportOptions } from '../export/xlsx-writer'

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
  /** 编辑事件(cell-change/edit-start/edit-commit;壳转 emit + 插件派发) */
  onEditEvent: (event: EditEventName, payload: unknown) => void
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

  // ---- 导出上下文(供 WorkbookExporter 取数) ----
  private workbook: WorkbookModel | null = null
  private activeIndex = 0
  private rendererOpts: RendererOptions = {}
  /** 下载默认文件名(壳可随 props 更新) */
  fileName: string | undefined = undefined
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
        this.renderer?.rebuildMetrics()
        this.refreshContentSize()
        this.render()
      },
      rebuildOverlays: () => {
        if (this.sheet && this.renderer) void this.overlays.build(this.sheet, this.renderer, this.view)
      },
      emit: (event, payload) => this.hooks.onEditEvent(event, payload),
    }
    this.edit = new EditController(editHost)
    this.editorHost = new CellEditorHost(els.editorSlot, (row, col) => this.rectOf(row, col))
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
    this.renderer = new CanvasRenderer(this.els.canvas, sheet, workbook, zoom, opts)
    this.hooks.onRenderer(this.renderer)
    this.view.zoom = zoom
    this.view.scrollX = 0
    this.view.scrollY = 0
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
  }

  /** 同步滚动量(壳的 scroll 事件里调) */
  setScroll(scrollX: number, scrollY: number): void {
    this.view.scrollX = scrollX
    this.view.scrollY = scrollY
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

  // ---- 维度 / 脏状态 命令式 API(E3.5) ----
  /** 程序化设列宽(px,模型单位/非缩放);editable 时走命令栈(可撤销+发 dim-change+记脏)。 */
  setColumnWidth(col: number, width: number): boolean {
    return this.edit.setDimension('col', col, width)
  }
  /** 程序化设行高(px,模型单位/非缩放);editable 时走命令栈。 */
  setRowHeight(row: number, height: number): boolean {
    return this.edit.setDimension('row', row, height)
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
    this.overlays.dispose()
    this.editorHost.dispose()
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
      this.dragMoved = true
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
    sc.style.cursor = r.cellHyperlink(cell.row, cell.col) ? 'pointer' : 'cell'
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
    const r = this.renderer
    const sc = this.els.scroller
    const c = this.selActive
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
    if (!r || !s) return
    // 防超大选区卡死: 复制范围软上限
    const rowEnd = Math.min(s.bottom, s.top + 4999)
    const colEnd = Math.min(s.right, s.left + 255)
    const lines: string[] = []
    const htmlRows: string[] = []
    for (let row = s.top; row <= rowEnd; row++) {
      const cells: string[] = []
      const htmlCells: string[] = []
      for (let col = s.left; col <= colEnd; col++) {
        const text = r.cellText(row, col)
        cells.push(text)
        const css = r.cellInlineStyle(row, col)
        htmlCells.push(`<td${css ? ` style="${css}"` : ''}>${escapeHtml(text)}</td>`)
      }
      lines.push(cells.join('\t'))
      htmlRows.push(`<tr>${htmlCells.join('')}</tr>`)
    }
    const tsv = lines.join('\n')
    const html = `<table border="1" style="border-collapse:collapse">${htmlRows.join('')}</table>`
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
    this.edit.refreshEngine() // recalc/formulaEngine 可能变了 → 重置引擎并按需点火
  }

  /** 该格当前是否可编辑(综合 editable + readOnlyRanges + cellReadOnly) */
  isCellEditable(row: number, col: number): boolean {
    return this.sheet ? resolveEditable(this.sheet, row, col, this.editCfg) : false
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

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
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
