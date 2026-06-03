/** 冻结窗格几何: 计算冻结区的像素尺寸，以及四象限分区。 */
import type { SheetModel } from '../model/types'
import type { GridMetrics } from './grid-metrics'

export interface FreezeGeometry {
  frozenRows: number
  frozenCols: number
  /** 冻结列总宽 px(网格坐标) */
  frozenWidth: number
  /** 冻结行总高 px */
  frozenHeight: number
}

export function computeFreeze(sheet: SheetModel, metrics: GridMetrics): FreezeGeometry {
  const frozenRows = Math.min(sheet.freeze.frozenRows, metrics.rows)
  const frozenCols = Math.min(sheet.freeze.frozenCols, metrics.cols)
  return {
    frozenRows,
    frozenCols,
    frozenWidth: metrics.colLeft(frozenCols),
    frozenHeight: metrics.rowTop(frozenRows),
  }
}
