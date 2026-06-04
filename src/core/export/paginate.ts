/**
 * 把一张(整表)底图按纸张可绘制区切片成多页(支持二维: 竖向 + 横向)。
 * - fitToWidth: 正文整体缩放到页宽(单列页),过高竖向跨页。
 * - 否则: 自然尺寸×zoom;过宽则横向跨页,过高则竖向跨页(像 Excel 页矩阵, 默认"先下后右")。
 * 打印标题: repeatTop(标题行×正文列,等宽 base)贴每页顶;repeatLeft(标题列×正文行,等高 base)贴每页左;
 * corner(标题列×标题行)贴左上角。切片在像素层进行(位图可跨行/列切,可接受)。
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
  /** 打印标题行条(标题行 × 正文列),贴每页顶部 */
  repeatTop?: { canvas: HTMLCanvasElement; heightCss: number }
  /** 打印标题列条(标题列 × 正文行),贴每页左侧 */
  repeatLeft?: { canvas: HTMLCanvasElement; widthCss: number }
  /** 标题角(标题列 × 标题行),贴左上角 */
  corner?: { canvas: HTMLCanvasElement }
}

interface Band {
  s: number
  len: number
}
/** 把 total 像素按 bandSize 切成若干段(末段取余)。(导出供测) */
export function makeBands(total: number, bandSize: number): Band[] {
  const out: Band[] = []
  if (bandSize >= total) return [{ s: 0, len: total }]
  for (let s = 0; s < total - 0.5; s += bandSize) {
    const s0 = Math.round(s)
    const len = Math.min(Math.round(bandSize), total - s0)
    if (len < 1) break
    out.push({ s: s0, len })
    if (out.length > 5000) break
  }
  return out
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

  // 标题条像素 / css 尺寸(同 base scale)
  const titleColPx = page.repeatLeft ? page.repeatLeft.canvas.width : 0
  const titleRowPx = page.repeatTop ? page.repeatTop.canvas.height : 0
  const titleColCss = page.repeatLeft ? page.repeatLeft.widthCss : 0
  const titleRowCss = page.repeatTop ? page.repeatTop.heightCss : 0

  // 缩放因子: fitToWidth → (正文+标题列)整体贴页宽;否则自然×zoom
  let factor: number
  if (page.fitToWidth) {
    const totalWcss = bodyWcss + titleColCss
    factor = totalWcss > 0 ? page.contentWmm / (totalWcss * MM_PER_PX) : 1
  } else {
    factor = zoom
  }
  const mmPerCss = MM_PER_PX * factor
  const titleColsWmm = titleColCss * mmPerCss
  const titleRowsHmm = titleRowCss * mmPerCss
  const availWmm = Math.max(1, page.contentWmm - titleColsWmm)
  const availHmm = Math.max(1, page.contentHmm - titleRowsHmm)
  const drawBodyWmm = naturalWmm * factor
  const drawBodyHmm = naturalHmm * factor
  if (drawBodyWmm < 0.01 || drawBodyHmm < 0.01) return []

  // 每页正文像素带(fitToWidth 横向不分页 → 整宽一带)
  const bandPxW = page.fitToWidth ? baseW : Math.max(1, (availWmm / drawBodyWmm) * baseW)
  const bandPxH = Math.max(1, (availHmm / drawBodyHmm) * baseH)
  const colBands = makeBands(baseW, bandPxW)
  const rowBands = makeBands(baseH, bandPxH)

  const pages: SlicedPage[] = []
  // 先下后右(Excel 默认 pageOrder=downThenOver)
  for (const cb of colBands) {
    for (const rb of rowBands) {
      const out = document.createElement('canvas')
      out.width = titleColPx + cb.len
      out.height = titleRowPx + rb.len
      const ctx = out.getContext('2d')
      if (!ctx) continue
      if (page.corner && titleColPx && titleRowPx) {
        ctx.drawImage(page.corner.canvas, 0, 0, titleColPx, titleRowPx, 0, 0, titleColPx, titleRowPx)
      }
      if (page.repeatTop && titleRowPx) {
        ctx.drawImage(page.repeatTop.canvas, cb.s, 0, cb.len, titleRowPx, titleColPx, 0, cb.len, titleRowPx)
      }
      if (page.repeatLeft && titleColPx) {
        ctx.drawImage(page.repeatLeft.canvas, 0, rb.s, titleColPx, rb.len, 0, titleRowPx, titleColPx, rb.len)
      }
      ctx.drawImage(base, cb.s, rb.s, cb.len, rb.len, titleColPx, titleRowPx, cb.len, rb.len)
      pages.push({
        canvas: out,
        widthMm: titleColsWmm + (cb.len / baseW) * drawBodyWmm,
        heightMm: titleRowsHmm + (rb.len / baseH) * drawBodyHmm,
      })
    }
  }
  return pages
}
