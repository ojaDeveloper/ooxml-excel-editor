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
import type { MergeRange, SheetModel, WorkbookModel } from '../model/types'
import { CanvasRenderer, type RendererOptions, type ViewState } from '../render/canvas-renderer'
import { OverlayManager, type OverlayQuads } from './overlay-manager'

export type Cell = { row: number; col: number }
export interface TooltipState {
  text: string
  x: number
  y: number
  kind: 'overflow' | 'comment'
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
  /** 点中自动筛选下拉按钮(壳打开 FilterPopup —— filter 仍在壳里) */
  onFilterButton: (col: number) => void
  /** 悬停提示变化(壳渲染 tooltip DOM) */
  onTooltip: (tip: TooltipState | null) => void
}

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

  // ---- 选区模型 ----
  private selAnchor: Cell | null = null // 固定角(扩选时不动)
  private selActive: Cell | null = null // 活动角(移动/扩选时变)
  private selMode: 'range' | 'rows' | 'cols' = 'range'

  // ---- 拖拽态 ----
  private dragMode: 'none' | 'cell' | 'row' | 'col' | 'resize-col' | 'resize-row' = 'none'
  private resizeTarget = -1 // 正在拖拽改宽高的列/行索引
  private resizeStartPos = 0 // 起始鼠标坐标(px)
  private resizeStartSize = 0 // 起始宽/高(px)
  private dragMoved = false

  constructor(
    private els: ViewerControllerEls,
    private hooks: ViewerControllerHooks,
  ) {
    this.overlays = new OverlayManager(els.overlays)
  }

  /** 切表/换簿/主题变化: 重建渲染器,重置滚动,量尺寸,建叠加层,绘制 */
  rebuild(sheet: SheetModel, workbook: WorkbookModel, zoom: number, opts: RendererOptions): void {
    this.sheet = sheet
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
  /** 双击列边界: 自适应列宽 */
  autoFitColumn(col: number): void {
    this.renderer?.autoFitColumn(col)
    this.refreshContentSize()
    this.render()
  }
  /** 双击行边界: 自适应行高 */
  autoFitRow(row: number): void {
    this.renderer?.autoFitRow(row)
    this.refreshContentSize()
    this.render()
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
        this.hooks.onFilterButton(fcol)
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
          return
        }
      } else if (p.x < r.metrics.rowHeaderWidth) {
        const b = this.nearRowBorder(p.x, p.y)
        if (b) {
          this.dragMode = 'resize-row'
          this.resizeTarget = b.row
          this.resizeStartPos = p.y
          this.resizeStartSize = r.metrics.rowHeight(b.row)
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
    }
    this.dragMode = 'none'
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
      // 非边界 → 双击单元格事件
      const cell = r.cellAtScreen(this.view, p.x, p.y)
      if (cell) this.hooks.onCellDblClick(cell.row, cell.col, r.cellText(cell.row, cell.col))
    }
  }

  // ====================== 键盘交互 ======================

  private pageRows(): number {
    const r = this.renderer
    if (!r) return 10
    return Math.max(1, Math.floor((this.view.height - r.metrics.colHeaderHeight) / r.defaultRowPx) - 1)
  }

  onKeyDown(e: KeyboardEvent): void {
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
