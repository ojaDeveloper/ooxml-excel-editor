/** 把 drawing 锚点换算成网格内容坐标(不含表头)的像素矩形。 */
import type { ImageAnchor } from '../model/types'
import type { GridMetrics } from '../layout/grid-metrics'
import { emuToPx } from '../layout/units'

export interface ContentRect {
  left: number
  top: number
  width: number
  height: number
}

export function anchorRect(metrics: GridMetrics, anchor: ImageAnchor): ContentRect {
  const z = metrics.zoom
  const px = (emu: number) => emuToPx(emu) * z // EMU 偏移也随缩放
  const left = metrics.colLeft(anchor.from.col) + px(anchor.from.colOffEmu)
  const top = metrics.rowTop(anchor.from.row) + px(anchor.from.rowOffEmu)

  let width: number
  let height: number
  if (anchor.to) {
    const right = metrics.colLeft(anchor.to.col) + px(anchor.to.colOffEmu)
    const bottom = metrics.rowTop(anchor.to.row) + px(anchor.to.rowOffEmu)
    width = Math.max(1, right - left)
    height = Math.max(1, bottom - top)
  } else {
    width = px(anchor.extWidthEmu ?? 0) || 120 * z
    height = px(anchor.extHeightEmu ?? 0) || 90 * z
  }
  return { left, top, width, height }
}
