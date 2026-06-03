/**
 * 把一张(整表)底图按纸张可绘制区竖向切片成多页。
 * fitToWidth 时内容缩放到页宽,过高自动跨页;否则按自然尺寸×zoom(96dpi)放置。
 * repeatTop(打印标题): 已渲染好的标题行条(同 scale、同列宽,故等宽于 base),贴在每页顶部重复。
 */
import { MM_PER_PX } from './raster'

export interface SlicedPage {
  /** 该页的图像切片(设备像素) */
  canvas: HTMLCanvasElement
  /** 贴到 PDF 页上的尺寸 mm */
  widthMm: number
  heightMm: number
}

export interface SliceOptions {
  contentWmm: number
  contentHmm: number
  fitToWidth: boolean
  /** 非 fitToWidth 时的打印缩放(pageSetup.scale/100),默认 1 */
  zoom?: number
  /** 打印标题: 每页顶部重复的标题行条 */
  repeatTop?: { canvas: HTMLCanvasElement; heightCss: number }
}

export function sliceToPages(
  base: HTMLCanvasElement,
  bodyWcss: number,
  bodyHcss: number,
  page: SliceOptions,
): SlicedPage[] {
  const baseW = base.width
  const baseH = base.height
  if (baseW < 1 || baseH < 1 || bodyWcss < 1 || bodyHcss < 1) return []

  const zoom = page.zoom && page.zoom > 0 ? page.zoom : 1
  const naturalWmm = bodyWcss * MM_PER_PX
  const naturalHmm = bodyHcss * MM_PER_PX
  // 缩放因子: fitToWidth 强制贴页宽;否则自然尺寸×zoom,但不超过页宽(本管线不横向分页)
  const drawWmm = page.fitToWidth ? page.contentWmm : Math.min(naturalWmm * zoom, page.contentWmm)
  const factor = naturalWmm > 0 ? drawWmm / naturalWmm : 1
  const drawHmmTotal = naturalHmm * factor
  if (drawHmmTotal < 0.01) return []

  // 标题条(同 scale,heightCss 是其 css 高;贴页时也按 factor 缩放)
  const strip = page.repeatTop
  const stripPx = strip ? strip.canvas.height : 0
  const stripHmm = strip ? strip.heightCss * MM_PER_PX * factor : 0
  const availHmm = page.contentHmm - stripHmm
  // 标题比整页还高 → 放弃重复标题,退化为无标题分页
  const useStrip = strip && availHmm > 1

  // 每页正文对应的底图像素高度(扣掉标题占用的页面高度)
  const bandPx = ((useStrip ? availHmm : page.contentHmm) / drawHmmTotal) * baseH
  if (!(bandPx > 0)) return []

  const pages: SlicedPage[] = []
  for (let sy = 0; sy < baseH - 0.5; sy += bandPx) {
    const sh = Math.min(Math.round(bandPx), baseH - Math.round(sy))
    if (sh < 1) break
    const bandHmm = (sh / baseH) * drawHmmTotal
    const out = document.createElement('canvas')
    out.width = baseW
    out.height = (useStrip ? stripPx : 0) + sh
    const octx = out.getContext('2d')
    if (!octx) break
    if (useStrip) octx.drawImage(strip!.canvas, 0, 0, strip!.canvas.width, stripPx, 0, 0, baseW, stripPx)
    octx.drawImage(base, 0, Math.round(sy), baseW, sh, 0, useStrip ? stripPx : 0, baseW, sh)
    pages.push({ canvas: out, widthMm: drawWmm, heightMm: (useStrip ? stripHmm : 0) + bandHmm })
    if (pages.length > 5000) break
  }
  return pages
}
