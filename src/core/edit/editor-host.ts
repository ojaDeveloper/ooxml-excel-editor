/**
 * CellEditorHost(框架无关)—— 单一活动编辑器的 DOM 宿主(结构同 PluginOverlayHost)。
 * 把工厂返回的 DOM 挂到独立 .editor-slot 层,按 rectOf 定位,随 renderTick 重定位;
 * 同时只允许一个活动编辑器。Vue / React 壳只给挂载节点 + rectOf,逻辑全在这里(两壳共用)。
 */
import type { CellEditorContext, CellEditorFactory, Rect } from './editor-context'

interface ActiveEditor {
  el: HTMLElement
  destroy?: () => void
  row: number
  col: number
  /** 可选定位覆盖:合并单元格编辑时返回整片合并区,而非单格 */
  rectOverride?: () => Rect | null
}

export class CellEditorHost {
  private active: ActiveEditor | null = null

  constructor(
    private container: HTMLElement,
    private rectOf: (row: number, col: number) => Rect | null,
  ) {}

  isActive(): boolean {
    return this.active !== null
  }
  activeCell(): { row: number; col: number } | null {
    return this.active ? { row: this.active.row, col: this.active.col } : null
  }

  /**
   * 挂载一个编辑器(先卸旧)。返回是否成功(工厂产出有效 DOM)。
   * rectOverride 给定时用它定位(合并单元格 → 整片合并区),否则按单格 rectOf。
   */
  mount(
    row: number,
    col: number,
    factory: CellEditorFactory,
    ctx: CellEditorContext,
    rectOverride?: () => Rect | null,
  ): boolean {
    this.unmount()
    let el: HTMLElement
    let destroy: (() => void) | undefined
    try {
      const made = factory(ctx)
      if (made instanceof HTMLElement) el = made
      else {
        el = made.el
        destroy = made.destroy
      }
    } catch (e) {
      console.warn('[ooxml-preview] 单元格编辑器工厂出错:', e)
      return false
    }
    el.style.position = 'absolute'
    el.style.boxSizing = 'border-box'
    this.container.appendChild(el)
    this.active = { el, destroy, row, col, rectOverride }
    this.position()
    return true
  }

  /** 按当前 rectOf 重定位活动编辑器(滚动/缩放后跟随) */
  position(): void {
    const a = this.active
    if (!a) return
    const r = a.rectOverride ? a.rectOverride() : this.rectOf(a.row, a.col)
    if (!r) return
    a.el.style.left = r.x + 'px'
    a.el.style.top = r.y + 'px'
    // 精确贴合单元格(像 WPS:编辑框正好盖住该格,不放大)。
    // 用 width/height 而非 min-*,否则 <input> 等控件会按自身固有尺寸(~20 字符)撑大。
    a.el.style.width = r.w + 'px'
    a.el.style.height = r.h + 'px'
  }

  /** 卸载活动编辑器 */
  unmount(): void {
    const a = this.active
    if (!a) return
    this.active = null
    try {
      a.destroy?.()
    } catch (e) {
      console.warn('[ooxml-preview] 编辑器 destroy 出错:', e)
    }
    a.el.remove()
  }

  dispose(): void {
    this.unmount()
  }
}
