/**
 * 插件 overlay 宿主(框架无关)。每个 tick 调插件的 overlay(ctx) 拿 DOM 节点,挂到容器里。
 * Vue / React 壳各自把自己的 overlay-slot 容器交给它,在重绘 tick 时调用 render —— 同一套逻辑。
 */
import type { ExcelPlugin, OverlayContext, OverlayNode } from '../plugin'

function toArray(node: OverlayNode): HTMLElement[] {
  if (!node) return []
  return Array.isArray(node) ? node.filter(Boolean) : [node]
}

export class PluginOverlayHost {
  constructor(private container: HTMLElement) {}

  /** 重渲所有插件的 overlay:依次调 overlay(ctx),把返回的 DOM 节点替换进容器。 */
  render(plugins: readonly ExcelPlugin[], ctx: OverlayContext): void {
    const next: HTMLElement[] = []
    for (const p of plugins) {
      if (!p.overlay) continue
      try {
        next.push(...toArray(p.overlay(ctx)))
      } catch (e) {
        console.warn('[ooxml-preview] 插件 overlay 渲染失败:', p.name, e)
      }
    }
    // 同引用就地复用(节点被 move,不重建),引用变了才换 —— replaceChildren 自动处理
    this.container.replaceChildren(...next)
  }

  dispose(): void {
    this.container.replaceChildren()
  }
}
