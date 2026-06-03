/**
 * 把图片/图表/形状合成到导出底图上。
 * 渲染器只画格子(填充/边框/条件格式/迷你图/文本);图片/图表/形状在屏幕上是 DOM 叠加层,
 * 导出时必须用同一套导出空间几何(zoom=1)算出位置,补画到底图 canvas 上。
 */
import type { ExportToCanvasResult } from '../render/canvas-renderer'
import type { ImageAnchor, ShapeSpec } from '../model/types'
import { anchorRect, type ContentRect } from '../overlay/anchor'

export interface ExportDecorations {
  /** 已加载的图片源(HTMLImageElement / 离屏 canvas 等) + 其锚点 */
  images?: { source: CanvasImageSource; anchor: ImageAnchor }[]
  /** 已栅格化的图表源(如 echarts getDataURL 加载成的 Image) + 锚点 */
  charts?: { source: CanvasImageSource; anchor: ImageAnchor }[]
  /** 形状/文本框(按 ShapeSpec 直接画) */
  shapes?: ShapeSpec[]
}

/** 把装饰合成到 base.canvas(原地修改)。 */
export function compositeOverlays(base: ExportToCanvasResult, deco?: ExportDecorations): void {
  if (!deco || (!deco.images?.length && !deco.charts?.length && !deco.shapes?.length)) return
  const ctx = base.canvas.getContext('2d')
  if (!ctx) return
  const s = base.scale

  // 网格内容坐标(zoom=1) → 底图设备像素
  const toDevice = (rect: ContentRect) => ({
    x: (base.originX + rect.left - base.gridOriginX) * s,
    y: (base.originY + rect.top - base.gridOriginY) * s,
    w: rect.width * s,
    h: rect.height * s,
  })

  ctx.save()
  for (const im of deco.images ?? []) {
    const r = toDevice(anchorRect(base.metrics, im.anchor))
    try {
      ctx.drawImage(im.source, r.x, r.y, r.w, r.h)
    } catch {
      /* 跨域污染等绘制失败,跳过该图 */
    }
  }
  for (const c of deco.charts ?? []) {
    const r = toDevice(anchorRect(base.metrics, c.anchor))
    try {
      ctx.drawImage(c.source, r.x, r.y, r.w, r.h)
    } catch {
      /* 跳过 */
    }
  }
  for (const sh of deco.shapes ?? []) {
    drawShape(ctx, sh, toDevice(anchorRect(base.metrics, sh.anchor)), s)
  }
  ctx.restore()
}

/** 画一个形状/文本框(近似屏幕样式: 填充 + 边框 + 居中文字) */
function drawShape(
  ctx: CanvasRenderingContext2D,
  sh: ShapeSpec,
  box: { x: number; y: number; w: number; h: number },
  scale: number,
): void {
  const { x, y, w, h } = box
  ctx.save()
  ctx.beginPath()
  if (sh.geom === 'ellipse') {
    ctx.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2)
  } else if (sh.geom === 'roundRect') {
    roundRectPath(ctx, x, y, w, h, 8 * scale)
  } else {
    ctx.rect(x, y, w, h)
  }
  if (sh.fillColor) {
    ctx.fillStyle = sh.fillColor
    ctx.fill()
  }
  if (sh.lineColor) {
    ctx.strokeStyle = sh.lineColor
    ctx.lineWidth = Math.max(1, scale)
    ctx.stroke()
  }
  if (sh.text) {
    ctx.clip()
    ctx.fillStyle = sh.textColor || '#000000'
    const fontPx = 11 * (96 / 72) * scale
    ctx.font = `${sh.bold ? 'bold ' : ''}${fontPx}px Calibri, sans-serif`
    ctx.textBaseline = 'middle'
    ctx.textAlign = sh.align === 'center' ? 'center' : sh.align === 'right' ? 'right' : 'left'
    const tx = sh.align === 'center' ? x + w / 2 : sh.align === 'right' ? x + w - 5 * scale : x + 5 * scale
    // 简单按换行拆,多行垂直居中
    const lines = sh.text.split('\n')
    const lineH = fontPx * 1.2
    let ty = y + h / 2 - ((lines.length - 1) * lineH) / 2
    for (const ln of lines) {
      ctx.fillText(ln, tx, ty)
      ty += lineH
    }
  }
  ctx.restore()
}

function roundRectPath(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, r: number): void {
  const rr = Math.min(r, w / 2, h / 2)
  ctx.moveTo(x + rr, y)
  ctx.arcTo(x + w, y, x + w, y + h, rr)
  ctx.arcTo(x + w, y + h, x, y + h, rr)
  ctx.arcTo(x, y + h, x, y, rr)
  ctx.arcTo(x, y, x + w, y, rr)
  ctx.closePath()
}
