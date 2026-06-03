/** 单元格填充绘制(纯色/图案/渐变)。 */
import type { Fill } from '../model/types'

export function paintFill(
  ctx: CanvasRenderingContext2D,
  fill: Fill,
  x: number,
  y: number,
  w: number,
  h: number,
): void {
  if (!fill || fill.type === 'none') return
  if (fill.type === 'solid') {
    if (!fill.fgColor) return
    ctx.fillStyle = fill.fgColor
    ctx.fillRect(x, y, w, h)
    return
  }
  if (fill.type === 'gradient') {
    const stops = fill.gradientStops || []
    if (!stops.length) return
    // 按角度确定渐变线(0°=左→右, 90°=上→下);缺省 90
    const angle = ((fill.gradientAngle ?? 90) * Math.PI) / 180
    const cx = x + w / 2
    const cy = y + h / 2
    const half = (Math.abs(w * Math.cos(angle)) + Math.abs(h * Math.sin(angle))) / 2
    const dx = Math.cos(angle) * half
    const dy = Math.sin(angle) * half
    const grad = ctx.createLinearGradient(cx - dx, cy - dy, cx + dx, cy + dy)
    for (const s of stops) grad.addColorStop(Math.min(1, Math.max(0, s.position)), s.color)
    ctx.fillStyle = grad
    ctx.fillRect(x, y, w, h)
    return
  }
  if (fill.type === 'pattern') {
    // 背景先铺 bgColor，再用 fgColor 画简化图案
    if (fill.bgColor) {
      ctx.fillStyle = fill.bgColor
      ctx.fillRect(x, y, w, h)
    }
    if (fill.fgColor) {
      paintPattern(ctx, fill.pattern || 'gray125', fill.fgColor, x, y, w, h)
    }
  }
}

function paintPattern(
  ctx: CanvasRenderingContext2D,
  pattern: string,
  color: string,
  x: number,
  y: number,
  w: number,
  h: number,
): void {
  ctx.save()
  ctx.beginPath()
  ctx.rect(x, y, w, h)
  ctx.clip()
  ctx.fillStyle = color
  ctx.strokeStyle = color
  // 灰度类图案近似成半透明铺底
  const density: Record<string, number> = {
    gray125: 0.12,
    gray0625: 0.06,
    lightGray: 0.25,
    mediumGray: 0.5,
    darkGray: 0.75,
    gray: 0.5,
  }
  if (pattern in density) {
    ctx.globalAlpha = density[pattern]
    ctx.fillRect(x, y, w, h)
  } else if (pattern.includes('Horizontal')) {
    ctx.lineWidth = 1
    for (let yy = y; yy < y + h; yy += 4) line(ctx, x, yy, x + w, yy)
  } else if (pattern.includes('Vertical')) {
    ctx.lineWidth = 1
    for (let xx = x; xx < x + w; xx += 4) line(ctx, xx, y, xx, y + h)
  } else if (pattern.includes('Grid') || pattern.includes('Trellis')) {
    for (let yy = y; yy < y + h; yy += 4) line(ctx, x, yy, x + w, yy)
    for (let xx = x; xx < x + w; xx += 4) line(ctx, xx, y, xx, y + h)
  } else {
    // 其它斜线图案
    ctx.lineWidth = 1
    for (let d = -h; d < w; d += 4) line(ctx, x + d, y + h, x + d + h, y)
  }
  ctx.restore()
}

function line(ctx: CanvasRenderingContext2D, x1: number, y1: number, x2: number, y2: number) {
  ctx.beginPath()
  ctx.moveTo(x1, y1)
  ctx.lineTo(x2, y2)
  ctx.stroke()
}
