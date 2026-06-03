/** 自动筛选下拉按钮(只画样式，预览不交互)。 */
import type { MergeRange } from '../model/types'

const BTN = 16

/** 判断某 cell 是否是自动筛选表头(范围首行) */
export function isFilterHeader(range: MergeRange | undefined, row: number, col: number): boolean {
  if (!range) return false
  return row === range.top && col >= range.left && col <= range.right
}

/** 在单元格右侧画一个筛选下拉按钮。返回按钮占用的宽度(供文本避让)。 */
export function drawFilterButton(
  ctx: CanvasRenderingContext2D,
  cellX: number,
  cellY: number,
  cellW: number,
  cellH: number,
): number {
  const size = Math.min(BTN, cellW - 2, cellH - 2)
  if (size < 8) return 0
  const x = cellX + cellW - size - 1
  const y = cellY + (cellH - size) / 2
  ctx.save()
  ctx.fillStyle = '#FFFFFF'
  ctx.strokeStyle = '#A0A0A0'
  ctx.lineWidth = 1
  ctx.beginPath()
  ctx.rect(Math.round(x) + 0.5, Math.round(y) + 0.5, size, size)
  ctx.fill()
  ctx.stroke()
  // 下三角
  ctx.fillStyle = '#5B5B5B'
  const cx = x + size / 2
  const cy = y + size / 2
  ctx.beginPath()
  ctx.moveTo(cx - 3, cy - 1.5)
  ctx.lineTo(cx + 3, cy - 1.5)
  ctx.lineTo(cx, cy + 2.5)
  ctx.closePath()
  ctx.fill()
  ctx.restore()
  return size + 2
}
