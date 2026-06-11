/**
 * 叠加层管理器(框架无关)。
 * 管理图片 / 图表(echarts)/ 形状文本框的 DOM 元素:按冻结四象限挂载、随滚动/缩放定位、销毁。
 * 纯 DOM + renderer 几何,零框架依赖 —— Vue / React 壳共用。
 */
import type * as EChartsNS from 'echarts'
import type { ChartSpec, ImageAnchor, SheetModel } from '../model/types'
import type { CanvasRenderer, ViewState } from '../render/canvas-renderer'
import { anchorRect } from '../overlay/anchor'
import { chartToOption } from '../overlay/chart-mapper'
import { loadECharts } from '../overlay/echarts-loader'

type Quad = 'main' | 'frow' | 'fcol' | 'corner'

/** 冻结四象限容器: 主区 / 冻结行 / 冻结列 / 冻结角 */
export interface OverlayQuads {
  main: HTMLElement
  frow: HTMLElement
  fcol: HTMLElement
  corner: HTMLElement
}

interface ChartInstance {
  el: HTMLDivElement
  inst: EChartsNS.ECharts
  spec: ChartSpec
  quad: Quad
  lastW: number
  lastH: number
}
interface ImageEl {
  el: HTMLImageElement
  anchorIdx: number
  quad: Quad
}
interface ChartPlaceholder {
  el: HTMLDivElement
  spec: ChartSpec
  quad: Quad
}
interface ShapeEl {
  el: HTMLDivElement
  shapeIdx: number
  quad: Quad
}
interface PivotButtonEl {
  el: HTMLButtonElement
  tableIdx: number
  buttonIdx: number
  quad: Quad
}

const CHART_BOX_CSS =
  'position:absolute;background:rgba(255,255,255,0.96);border:1px solid #e2e4e7;box-shadow:0 1px 4px rgba(0,0,0,0.08);'
const PLACEHOLDER_CSS =
  'display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;color:#6b7785;border-style:dashed;box-sizing:border-box;padding:8px;overflow:hidden;'

export class OverlayManager {
  private chartInstances: ChartInstance[] = []
  private imageEls: ImageEl[] = []
  private chartPlaceholders: ChartPlaceholder[] = []
  private shapeEls: ShapeEl[] = []
  private pivotButtonEls: PivotButtonEl[] = []
  /** 防止异步加载 echarts 期间已切表/重建: build 自增,await 后 token 变了就放弃 */
  private buildToken = 0

  constructor(private quads: OverlayQuads) {}

  private container(q: Quad): HTMLElement {
    return this.quads[q]
  }

  /** 锚点落在哪个象限(冻结行/列内的钉住,其余随滚动) */
  private quadrantOf(anchor: ImageAnchor, renderer: CanvasRenderer): Quad {
    const fz = renderer.freezeGeometry
    const top = anchor.from.row < fz.frozenRows
    const left = anchor.from.col < fz.frozenCols
    if (top && left) return 'corner'
    if (top) return 'frow'
    if (left) return 'fcol'
    return 'main'
  }

  dispose(): void {
    for (const c of this.chartInstances) {
      c.inst.dispose()
      c.el.remove()
    }
    this.chartInstances = []
    for (const p of this.chartPlaceholders) p.el.remove()
    this.chartPlaceholders = []
    for (const sh of this.shapeEls) sh.el.remove()
    this.shapeEls = []
    for (const p of this.pivotButtonEls) p.el.remove()
    this.pivotButtonEls = []
    for (const im of this.imageEls) im.el.remove()
    this.imageEls = []
  }

  /** 重建该表全部叠加层(先销毁再挂载),末尾定位一次 */
  async build(sheet: SheetModel, renderer: CanvasRenderer, view: ViewState): Promise<void> {
    this.dispose()
    const token = ++this.buildToken

    for (let i = 0; i < sheet.images.length; i++) {
      const img = sheet.images[i]
      const quad = this.quadrantOf(img, renderer)
      const el = document.createElement('img')
      el.src = img.src
      el.draggable = false
      el.style.position = 'absolute'
      el.style.objectFit = 'fill'
      el.style.pointerEvents = 'none' // 穿透到 canvas:点击/拖拽由控制器在 onMouseUp 经 imageHitAt 命中处理
      this.container(quad).appendChild(el)
      this.imageEls.push({ el, anchorIdx: i, quad })
    }

    // 形状 / 文本框
    for (let i = 0; i < sheet.shapes.length; i++) {
      const shape = sheet.shapes[i]
      const quad = this.quadrantOf(shape.anchor, renderer)
      const el = document.createElement('div')
      el.style.position = 'absolute'
      el.style.boxSizing = 'border-box'
      el.style.pointerEvents = 'none'
      el.style.overflow = 'hidden'
      el.style.display = 'flex'
      el.style.padding = '3px 5px'
      el.style.alignItems = 'center'
      el.style.justifyContent = shape.align === 'center' ? 'center' : shape.align === 'right' ? 'flex-end' : 'flex-start'
      el.style.whiteSpace = 'pre-wrap'
      el.style.wordBreak = 'break-word'
      el.style.lineHeight = '1.2'
      if (shape.fillColor) el.style.background = shape.fillColor
      if (shape.lineColor) el.style.border = `1px solid ${shape.lineColor}`
      if (shape.geom === 'roundRect') el.style.borderRadius = '8px'
      else if (shape.geom === 'ellipse') el.style.borderRadius = '50%'
      if (shape.textColor) el.style.color = shape.textColor
      if (shape.bold) el.style.fontWeight = 'bold'
      el.style.textAlign = shape.align ?? 'left'
      el.textContent = shape.text ?? ''
      this.container(quad).appendChild(el)
      this.shapeEls.push({ el, shapeIdx: i, quad })
    }

    if (sheet.charts.length) {
      let echarts: typeof EChartsNS | null = null
      try {
        echarts = await loadECharts()
      } catch (e) {
        // echarts 是可选 peer 依赖,宿主未安装时降级
        console.warn('[ooxml-preview] 未能加载 echarts,图表降级为占位提示。请安装依赖: npm i echarts', e)
      }
      // 异步加载期间可能已切表/重建 → token 变了就放弃
      if (token !== this.buildToken) return

      for (const chart of sheet.charts) {
        const quad = this.quadrantOf(chart.anchor, renderer)
        if (echarts) {
          const el = document.createElement('div')
          el.style.cssText = CHART_BOX_CSS
          this.container(quad).appendChild(el)
          const inst = echarts.init(el)
          try {
            inst.setOption(chartToOption(chart))
          } catch (e) {
            console.warn('[ooxml-preview] 图表渲染失败:', e)
          }
          this.chartInstances.push({ el, inst, spec: chart, quad, lastW: 0, lastH: 0 })
        } else {
          const el = document.createElement('div')
          el.style.cssText = CHART_BOX_CSS + PLACEHOLDER_CSS
          el.innerHTML =
            '<div style="font-size:22px;line-height:1">📊</div>' +
            '<div style="margin-top:6px;font-weight:600">图表</div>' +
            '<div style="margin-top:2px;font-size:11px;color:#9aa4ae">渲染需安装 echarts 依赖</div>'
          this.container(quad).appendChild(el)
          this.chartPlaceholders.push({ el, spec: chart, quad })
        }
      }
    }

    for (let ti = 0; ti < sheet.pivotTables.length; ti++) {
      const table = sheet.pivotTables[ti]
      for (let bi = 0; bi < table.buttons.length; bi++) {
        const btn = table.buttons[bi]
        const anchor: ImageAnchor = {
          src: '',
          from: { row: btn.row, col: btn.col, rowOffEmu: 0, colOffEmu: 0 },
          to: { row: btn.row + 1, col: btn.col + 1, rowOffEmu: 0, colOffEmu: 0 },
        }
        const quad = this.quadrantOf(anchor, renderer)
        const el = document.createElement('button')
        el.type = 'button'
        el.className = `ooxml-pivot-button ${btn.kind}`
        el.title = `${table.name}: ${btn.label}`
        el.setAttribute('aria-label', `透视表字段 ${btn.label}`)
        const label = document.createElement('span')
        label.className = 'ooxml-pivot-label'
        label.textContent = btn.label
        const caret = document.createElement('span')
        caret.className = 'ooxml-pivot-caret'
        caret.textContent = '▾'
        el.append(label, caret)
        this.container(quad).appendChild(el)
        this.pivotButtonEls.push({ el, tableIdx: ti, buttonIdx: bi, quad })
      }
    }
    this.position(sheet, renderer, view)
  }

  /** 按当前视图(滚动/缩放)重新定位四象限容器与各叠加层 */
  position(sheet: SheetModel, renderer: CanvasRenderer, view: ViewState): void {
    const hw = renderer.metrics.rowHeaderWidth
    const hh = renderer.metrics.colHeaderHeight
    const fz = renderer.freezeGeometry
    const fw = fz.frozenWidth
    const fh = fz.frozenHeight
    const bodyW = Math.max(0, view.width - hw - fw)
    const bodyH = Math.max(0, view.height - hh - fh)

    setQuad(this.quads.corner, hw, hh, fw, fh)
    setQuad(this.quads.frow, hw + fw, hh, bodyW, fh)
    setQuad(this.quads.fcol, hw, hh + fh, fw, bodyH)
    setQuad(this.quads.main, hw + fw, hh + fh, bodyW, bodyH)

    for (const im of this.imageEls) {
      const anchor = sheet.images[im.anchorIdx]
      if (!anchor) continue // 图片在 build 与 position 之间被移除(如转内嵌图)→ 跳过,待重建
      placeInQuad(im.el, anchorRect(renderer.metrics, anchor), im.quad, fw, fh, view)
    }
    const shapeFont = Math.round(11 * (96 / 72) * renderer.metrics.zoom)
    for (const sh of this.shapeEls) {
      sh.el.style.fontSize = shapeFont + 'px'
      placeInQuad(sh.el, anchorRect(renderer.metrics, sheet.shapes[sh.shapeIdx].anchor), sh.quad, fw, fh, view)
    }
    for (const c of this.chartInstances) {
      const rect = anchorRect(renderer.metrics, c.spec.anchor)
      placeInQuad(c.el, rect, c.quad, fw, fh, view)
      // 只有尺寸真变(缩放)才 resize;滚动只改位置,不必 resize(很费)
      if (c.lastW !== rect.width || c.lastH !== rect.height) {
        c.inst.resize()
        c.lastW = rect.width
        c.lastH = rect.height
      }
    }
    for (const p of this.chartPlaceholders) {
      placeInQuad(p.el, anchorRect(renderer.metrics, p.spec.anchor), p.quad, fw, fh, view)
    }
    for (const p of this.pivotButtonEls) {
      const btn = sheet.pivotTables[p.tableIdx]?.buttons[p.buttonIdx]
      if (!btn) continue
      const rect = cellRect(renderer, btn.row, btn.col)
      placeInQuad(p.el, rect, p.quad, fw, fh, view)
    }
  }
}

function cellRect(renderer: CanvasRenderer, row: number, col: number) {
  return {
    left: renderer.metrics.colLeft(col),
    top: renderer.metrics.rowTop(row),
    width: renderer.metrics.colWidth(col),
    height: renderer.metrics.rowHeight(row),
  }
}

function setQuad(el: HTMLElement, left: number, top: number, w: number, h: number) {
  el.style.left = left + 'px'
  el.style.top = top + 'px'
  el.style.width = Math.max(0, w) + 'px'
  el.style.height = Math.max(0, h) + 'px'
}

/** 把元素定位到所在象限容器内的相对坐标(冻结方向不减滚动量) */
function placeInQuad(
  el: HTMLElement,
  rect: { left: number; top: number; width: number; height: number },
  quad: Quad,
  fw: number,
  fh: number,
  view: ViewState,
) {
  const x = quad === 'main' || quad === 'frow' ? rect.left - fw - view.scrollX : rect.left
  const y = quad === 'main' || quad === 'fcol' ? rect.top - fh - view.scrollY : rect.top
  el.style.left = x + 'px'
  el.style.top = y + 'px'
  el.style.width = rect.width + 'px'
  el.style.height = rect.height + 'px'
}
