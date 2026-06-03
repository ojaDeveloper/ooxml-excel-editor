/** 单元格边框绘制。 */
import type { BorderEdge, BorderStyle } from '../model/types'

/** 返回 [线宽, 虚线数组, 是否双线] */
function borderSpec(style: BorderStyle): { width: number; dash: number[]; double: boolean } {
  switch (style) {
    case 'hair': return { width: 0.5, dash: [], double: false }
    case 'thin': return { width: 1, dash: [], double: false }
    case 'medium': return { width: 2, dash: [], double: false }
    case 'thick': return { width: 3, dash: [], double: false }
    case 'double': return { width: 1, dash: [], double: true }
    case 'dotted': return { width: 1, dash: [1, 1], double: false }
    case 'dashed': return { width: 1, dash: [3, 2], double: false }
    case 'mediumDashed': return { width: 2, dash: [4, 2], double: false }
    case 'dashDot': return { width: 1, dash: [3, 2, 1, 2], double: false }
    case 'mediumDashDot': return { width: 2, dash: [4, 2, 1, 2], double: false }
    case 'dashDotDot': return { width: 1, dash: [3, 2, 1, 2, 1, 2], double: false }
    case 'mediumDashDotDot': return { width: 2, dash: [4, 2, 1, 2, 1, 2], double: false }
    case 'slantDashDot': return { width: 1, dash: [4, 2], double: false }
    default: return { width: 1, dash: [], double: false }
  }
}

/**
 * 画一条边。(x1,y1)-(x2,y2) 为边的两端点(已是屏幕坐标)。
 * 横边传 y1===y2，竖边传 x1===x2。
 */
export function drawEdge(
  ctx: CanvasRenderingContext2D,
  edge: BorderEdge | undefined,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
): void {
  if (!edge || edge.style === 'none') return
  const spec = borderSpec(edge.style)
  ctx.save()
  ctx.strokeStyle = edge.color
  ctx.lineWidth = spec.width
  ctx.setLineDash(spec.dash)
  const horizontal = y1 === y2
  // 对齐到像素中线，避免模糊
  const align = (v: number) => Math.round(v) + (spec.width % 2 === 1 ? 0.5 : 0)
  if (spec.double) {
    const off = 1
    if (horizontal) {
      line(ctx, x1, align(y1 - off), x2, align(y1 - off))
      line(ctx, x1, align(y1 + off), x2, align(y1 + off))
    } else {
      line(ctx, align(x1 - off), y1, align(x1 - off), y2)
      line(ctx, align(x1 + off), y1, align(x1 + off), y2)
    }
  } else {
    if (horizontal) line(ctx, x1, align(y1), x2, align(y2))
    else line(ctx, align(x1), y1, align(x2), y2)
  }
  ctx.restore()
}

function line(ctx: CanvasRenderingContext2D, x1: number, y1: number, x2: number, y2: number) {
  ctx.beginPath()
  ctx.moveTo(x1, y1)
  ctx.lineTo(x2, y2)
  ctx.stroke()
}
