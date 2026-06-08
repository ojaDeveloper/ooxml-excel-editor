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
  /** Phase 1 长文本撑高 (2026-06-08): 给宽度返期望高度. host 取 max(cell h, 此值) 作为最终高度. */
  getDesiredHeight?: (widthPx: number) => number
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
    let getDesiredHeight: ((w: number) => number) | undefined
    try {
      // Phase 1 长文本撑高 (2026-06-08): 注入 reposition 给 ctx, 编辑器在 input 事件后调它即可重撑高
      ctx.reposition = () => this.position()
      const made = factory(ctx)
      if (made instanceof HTMLElement) el = made
      else {
        el = made.el
        destroy = made.destroy
        getDesiredHeight = made.getDesiredHeight
      }
    } catch (e) {
      console.warn('[ooxml-preview] 单元格编辑器工厂出错:', e)
      return false
    }
    el.style.position = 'absolute'
    el.style.boxSizing = 'border-box'
    this.container.appendChild(el)
    this.active = { el, destroy, row, col, rectOverride, getDesiredHeight }
    this.position()
    return true
  }

  /**
   * 按当前 rectOf 重定位活动编辑器(滚动/缩放后跟随).
   * Phase 1 长文本撑高 (2026-06-08): 若编辑器实现了 `getDesiredHeight`,
   * 最终高度 = max(单元格原高, 期望高度), 上限 viewport 一半防撑爆.
   * 宽度仍 = 列宽 (跟 WPS 一致, 仅向下溢出).
   */
  position(): void {
    const a = this.active
    if (!a) return
    const r = a.rectOverride ? a.rectOverride() : this.rectOf(a.row, a.col)
    if (!r) return
    a.el.style.left = r.x + 'px'
    a.el.style.top = r.y + 'px'
    // 用 width/height 而非 min-*,否则 <input> 等控件会按自身固有尺寸(~20 字符)撑大。
    a.el.style.width = r.w + 'px'
    let h = r.h
    if (a.getDesiredHeight) {
      try {
        const desired = a.getDesiredHeight(r.w)
        if (desired > h) {
          // 上限: viewport 高的一半, 不让编辑器把整个屏幕撑爆 (textarea 自己内部 overflow:auto 滚)
          const win = this.container.ownerDocument.defaultView
          const cap = Math.max(120, (win?.innerHeight ?? 600) * 0.5)
          h = Math.min(desired, cap)
        }
      } catch (e) {
        console.warn('[ooxml-preview] 编辑器 getDesiredHeight 出错:', e)
      }
    }
    a.el.style.height = h + 'px'
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
