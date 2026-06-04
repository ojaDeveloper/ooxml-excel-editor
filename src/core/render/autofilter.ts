/** 自动筛选下拉按钮(只画样式，预览不交互)。 */
import type { MergeRange } from '../model/types'

const BTN = 16

/** 判断某 cell 是否是自动筛选表头(范围首行) */
export function isFilterHeader(range: MergeRange | undefined, row: number, col: number): boolean {
  if (!range) return false
  return row === range.top && col >= range.left && col <= range.right
}

/** 单元格右侧筛选下拉按钮的方形区域(像素);供命中检测与绘制共用 */
export function filterButtonBox(
  cellX: number,
  cellY: number,
  cellW: number,
  cellH: number,
): { x: number; y: number; size: number } | null {
  const size = Math.min(BTN, cellW - 2, cellH - 2)
  if (size < 8) return null
  return { x: cellX + cellW - size - 1, y: cellY + (cellH - size) / 2, size }
}

/**
 * 在单元格右侧画一个筛选下拉按钮。active=true 时画蓝色漏斗(已筛选)。
 * 返回按钮占用的宽度(供文本避让)。
 */
export function drawFilterButton(
  ctx: CanvasRenderingContext2D,
  cellX: number,
  cellY: number,
  cellW: number,
  cellH: number,
  active = false,
): number {
  const box = filterButtonBox(cellX, cellY, cellW, cellH)
  if (!box) return 0
  const { x, y, size } = box
  ctx.save()
  ctx.fillStyle = active ? '#E3F0FF' : '#FFFFFF'
  ctx.strokeStyle = active ? '#1A73E8' : '#A0A0A0'
  ctx.lineWidth = 1
  ctx.beginPath()
  ctx.rect(Math.round(x) + 0.5, Math.round(y) + 0.5, size, size)
  ctx.fill()
  ctx.stroke()
  const cx = x + size / 2
  const cy = y + size / 2
  if (active) {
    // 漏斗(已筛选)
    ctx.fillStyle = '#1A73E8'
    ctx.beginPath()
    ctx.moveTo(cx - 3.5, cy - 3)
    ctx.lineTo(cx + 3.5, cy - 3)
    ctx.lineTo(cx + 1, cy + 0.5)
    ctx.lineTo(cx + 1, cy + 3.5)
    ctx.lineTo(cx - 1, cy + 2.5)
    ctx.lineTo(cx - 1, cy + 0.5)
    ctx.closePath()
    ctx.fill()
  } else {
    // 下三角
    ctx.fillStyle = '#5B5B5B'
    ctx.beginPath()
    ctx.moveTo(cx - 3, cy - 1.5)
    ctx.lineTo(cx + 3, cy - 1.5)
    ctx.lineTo(cx, cy + 2.5)
    ctx.closePath()
    ctx.fill()
  }
  ctx.restore()
  return size + 2
}
