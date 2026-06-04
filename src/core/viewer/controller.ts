/**
 * ViewerController(框架无关)—— 接管"交互式网格"的渲染引擎,供 Vue / React 壳共用。
 *
 * 本阶段(A2a)职责: renderer 生命周期、view 状态(滚动/缩放/尺寸)、render 调度(rAF)、
 * measure、spacer 尺寸、列宽行高拖拽/自适应、叠加层(OverlayManager)、几何 API(rectOf)。
 * 选区/交互/find/filter 仍在壳里,通过本控制器的 renderer/view + 方法操作。
 *
 * 与框架的桥接全走 hooks: getSelection(渲染时取当前选区)、onRenderer(把 renderer 镜像回壳的响应式)、
 * onRenderTick(壳据此重算 overlay slot 位置)。壳不需镜像 contentSize —— 控制器直接量 spacer。
 */
import type { MergeRange, SheetModel, WorkbookModel } from '../model/types'
import { CanvasRenderer, type RendererOptions, type ViewState } from '../render/canvas-renderer'
import { OverlayManager, type OverlayQuads } from './overlay-manager'

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
  /** 渲染时取当前选区(选区仍由壳持有) */
  getSelection: () => MergeRange | null
  /** renderer 重建时回调,壳据此镜像到响应式(保持现有 renderer.value 读法 + chrome 响应) */
  onRenderer: (renderer: CanvasRenderer | null) => void
  /** 每次绘制后回调,壳据此 +1 让 overlay slot 重算位置 */
  onRenderTick: () => void
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
    r.setSelection(this.hooks.getSelection())
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
}
