/**
 * 把一张(整表)底图按纸张可绘制区竖向切片成多页。
 * fitToWidth 时内容缩放到页宽,过高自动跨页;否则按自然尺寸(96dpi)放置。
 */
import { MM_PER_PX } from './raster'

export interface SlicedPage {
  /** 该页的图像切片(设备像素) */
  canvas: HTMLCanvasElement
  /** 贴到 PDF 页上的尺寸 mm */
  widthMm: number
  heightMm: number
}

export function sliceToPages(
  base: HTMLCanvasElement,
  bodyWcss: number,
  bodyHcss: number,
  page: { contentWmm: number; contentHmm: number; fitToWidth: boolean },
): SlicedPage[] {
  const baseW = base.width
  const baseH = base.height
  if (baseW < 1 || baseH < 1 || bodyWcss < 1 || bodyHcss < 1) return []

  const naturalWmm = bodyWcss * MM_PER_PX
  const naturalHmm = bodyHcss * MM_PER_PX
  // 缩放因子: fitToWidth 强制贴页宽(放大/缩小都可);否则不超过页宽
  const drawWmm = page.fitToWidth ? page.contentWmm : Math.min(naturalWmm, page.contentWmm)
  const factor = naturalWmm > 0 ? drawWmm / naturalWmm : 1
  const drawHmmTotal = naturalHmm * factor
  if (drawHmmTotal < 0.01) return []

  // 每页对应的底图像素高度
  const bandPx = (page.contentHmm / drawHmmTotal) * baseH
  if (!(bandPx > 0)) return []

  const pages: SlicedPage[] = []
  for (let sy = 0; sy < baseH - 0.5; sy += bandPx) {
    const sh = Math.min(Math.round(bandPx), baseH - Math.round(sy))
    if (sh < 1) break
    const slice = document.createElement('canvas')
    slice.width = baseW
    slice.height = sh
    const sctx = slice.getContext('2d')
    if (!sctx) break
    sctx.drawImage(base, 0, Math.round(sy), baseW, sh, 0, 0, baseW, sh)
    pages.push({ canvas: slice, widthMm: drawWmm, heightMm: (sh / baseH) * drawHmmTotal })
    // 防御: 极端情况下避免死循环
    if (pages.length > 5000) break
  }
  return pages
}
