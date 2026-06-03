/**
 * 视口/分区计算: 把"表头 + 冻结区 + 可滚动区"组织成绘制用的 pane 列表。
 * 每个 pane 给出: 屏幕裁剪矩形 + 该 pane 内容的滚动偏移 + 行列区间。
 */
import type { GridMetrics } from './grid-metrics'
import type { FreezeGeometry } from './freeze'

export interface Pane {
  /** 屏幕坐标裁剪区(含表头偏移之后) */
  clipX: number
  clipY: number
  clipW: number
  clipH: number
  /** 该 pane 绘制内容时减去的滚动量(冻结区为 0) */
  offsetX: number
  offsetY: number
  /** 行列区间 [startCol,endCol] / [startRow,endRow] */
  colStart: number
  colEnd: number
  rowStart: number
  rowEnd: number
}

export interface ViewportLayout {
  panes: Pane[]
  rowHeaderWidth: number
  colHeaderHeight: number
}

/**
 * @param scrollX 可滚动区横向滚动量
 * @param scrollY 可滚动区纵向滚动量
 * @param viewW   画布 css 宽
 * @param viewH   画布 css 高
 */
export function computeViewport(
  metrics: GridMetrics,
  freeze: FreezeGeometry,
  scrollX: number,
  scrollY: number,
  viewW: number,
  viewH: number,
): ViewportLayout {
  const hw = metrics.rowHeaderWidth
  const hh = metrics.colHeaderHeight
  const fw = freeze.frozenWidth
  const fh = freeze.frozenHeight

  // 主可滚动区的起点(屏幕坐标)
  const bodyX = hw + fw
  const bodyY = hh + fh
  const bodyW = Math.max(0, viewW - bodyX)
  const bodyH = Math.max(0, viewH - bodyY)

  // 各方向行列区间
  const [scStart, scEnd] = metrics.visibleColRange(scrollX + fw, bodyW) // 滚动区列(跳过冻结列)
  const [srStart, srEnd] = metrics.visibleRowRange(scrollY + fh, bodyH)

  const panes: Pane[] = []

  // 主滚动区(右下)
  panes.push({
    clipX: bodyX, clipY: bodyY, clipW: bodyW, clipH: bodyH,
    offsetX: scrollX, offsetY: scrollY,
    colStart: Math.max(scStart, freeze.frozenCols),
    colEnd: scEnd,
    rowStart: Math.max(srStart, freeze.frozenRows),
    rowEnd: srEnd,
  })

  // 冻结列(左下): 列固定，行随纵向滚动
  if (freeze.frozenCols > 0) {
    panes.push({
      clipX: hw, clipY: bodyY, clipW: fw, clipH: bodyH,
      offsetX: 0, offsetY: scrollY,
      colStart: 0, colEnd: freeze.frozenCols - 1,
      rowStart: Math.max(srStart, freeze.frozenRows), rowEnd: srEnd,
    })
  }

  // 冻结行(右上): 行固定，列随横向滚动
  if (freeze.frozenRows > 0) {
    panes.push({
      clipX: bodyX, clipY: hh, clipW: bodyW, clipH: fh,
      offsetX: scrollX, offsetY: 0,
      colStart: Math.max(scStart, freeze.frozenCols), colEnd: scEnd,
      rowStart: 0, rowEnd: freeze.frozenRows - 1,
    })
  }

  // 冻结角(左上): 完全固定
  if (freeze.frozenRows > 0 && freeze.frozenCols > 0) {
    panes.push({
      clipX: hw, clipY: hh, clipW: fw, clipH: fh,
      offsetX: 0, offsetY: 0,
      colStart: 0, colEnd: freeze.frozenCols - 1,
      rowStart: 0, rowEnd: freeze.frozenRows - 1,
    })
  }

  return { panes, rowHeaderWidth: hw, colHeaderHeight: hh }
}
