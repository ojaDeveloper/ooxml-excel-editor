/**
 * 网格几何: 列宽/行高换算后的累计偏移(前缀和) + 坐标↔行列二分查找。
 * 渲染器靠它把行列号映射到像素矩形，把滚动位置映射到可视行列区间。
 */
import type { SheetModel } from '../model/types'

/** Excel/WPS 网格上限:行 1,048,576 / 列 16,384(虚拟外推封顶,防失控) */
export const MAX_GRID_ROWS = 1048576
export const MAX_GRID_COLS = 16384

export class GridMetrics {
  readonly cols: number
  readonly rows: number
  /**
   * 虚拟行/列上限(≥ dimension,≤ Excel 上限)。**仅**用于"滚动出空行"的 spacer 尺寸 /
   * 可视区范围 / 命中夹取 / 表头;**不**改 totalWidth/Height(导出与 data-access 仍按 dimension)。
   */
  readonly vRows: number
  readonly vCols: number
  /** 行表头(显示行号)宽度 px */
  readonly rowHeaderWidth: number
  /** 列表头(显示列字母)高度 px */
  readonly colHeaderHeight: number

  private colLefts: number[] // 长度 cols+1，colLefts[c] = 第 c 列左边缘(相对网格原点)
  private rowTops: number[] // 长度 rows+1
  private colWidths: number[]
  private rowHeights: number[]

  /** 当前缩放比例(几何与字体同步按它缩放，保证缩放后排版一致) */
  readonly zoom: number

  constructor(private sheet: SheetModel, zoom = 1, virtualRows = 0, virtualCols = 0) {
    this.zoom = zoom
    this.cols = Math.max(sheet.dimension.cols, 1)
    this.rows = Math.max(sheet.dimension.rows, 1)
    this.vRows = Math.min(Math.max(this.rows, virtualRows), MAX_GRID_ROWS)
    this.vCols = Math.min(Math.max(this.cols, virtualCols), MAX_GRID_COLS)
    this.colHeaderHeight = Math.round(22 * zoom)
    // 行表头宽度随最大行号位数自适应(按虚拟行数,留足位数)
    this.rowHeaderWidth = Math.round(Math.max(40, String(this.vRows).length * 8 + 16) * zoom)

    this.colWidths = new Array(this.cols)
    this.colLefts = new Array(this.cols + 1)
    this.colLefts[0] = 0
    for (let c = 0; c < this.cols; c++) {
      const info = sheet.columns.get(c)
      const w = (info?.hidden ? 0 : info?.width ?? sheet.defaultColWidth) * zoom
      this.colWidths[c] = w
      this.colLefts[c + 1] = this.colLefts[c] + w
    }

    this.rowHeights = new Array(this.rows)
    this.rowTops = new Array(this.rows + 1)
    this.rowTops[0] = 0
    for (let r = 0; r < this.rows; r++) {
      const info = sheet.rows.get(r)
      const h = (info?.hidden ? 0 : info?.height ?? sheet.defaultRowHeight) * zoom
      this.rowHeights[r] = h
      this.rowTops[r + 1] = this.rowTops[r] + h
    }
  }

  colWidth(c: number): number {
    return c >= 0 && c < this.cols ? this.colWidths[c] : this.sheet.defaultColWidth * this.zoom
  }
  rowHeight(r: number): number {
    return r >= 0 && r < this.rows ? this.rowHeights[r] : this.sheet.defaultRowHeight * this.zoom
  }
  private get dcw(): number {
    return this.sheet.defaultColWidth * this.zoom
  }
  private get drh(): number {
    return this.sheet.defaultRowHeight * this.zoom
  }

  /** 第 c 列左边缘(网格坐标，不含表头)。超出数据范围按默认列宽外推(模拟 Excel 无限网格)。 */
  colLeft(c: number): number {
    if (c <= 0) return 0
    if (c >= this.cols) return this.colLefts[this.cols] + (c - this.cols) * this.dcw
    return this.colLefts[c]
  }
  rowTop(r: number): number {
    if (r <= 0) return 0
    if (r >= this.rows) return this.rowTops[this.rows] + (r - this.rows) * this.drh
    return this.rowTops[r]
  }
  get totalWidth(): number {
    return this.colLefts[this.cols]
  }
  get totalHeight(): number {
    return this.rowTops[this.rows]
  }
  /** 含虚拟外推的总宽/高(spacer 尺寸用;= totalWidth + 虚拟列外推) */
  get virtualWidth(): number {
    return this.colLeft(this.vCols)
  }
  get virtualHeight(): number {
    return this.rowTop(this.vRows)
  }

  /** 给定网格 x 坐标，返回所在列。超出数据范围按默认列宽外推。 */
  colAt(x: number): number {
    if (x >= this.totalWidth) {
      return this.cols + Math.floor((x - this.totalWidth) / this.dcw)
    }
    return clampSearch(this.colLefts, x, this.cols)
  }
  rowAt(y: number): number {
    if (y >= this.totalHeight) {
      return this.rows + Math.floor((y - this.totalHeight) / this.drh)
    }
    return clampSearch(this.rowTops, y, this.rows)
  }

  /** 数据单元格的可视区列区间(end 夹到虚拟范围 vCols-1,允许滚动出空列;空格 paint 为 no-op) */
  visibleColRange(scrollX: number, viewW: number): [number, number] {
    const start = this.colAt(scrollX)
    const end = this.colAt(scrollX + viewW)
    return [Math.min(start, this.vCols - 1), Math.min(end + 1, this.vCols - 1)]
  }
  visibleRowRange(scrollY: number, viewH: number): [number, number] {
    const start = this.rowAt(scrollY)
    const end = this.rowAt(scrollY + viewH)
    return [Math.min(start, this.vRows - 1), Math.min(end + 1, this.vRows - 1)]
  }

  /** 网格线/表头的可视区列区间(可超出数据范围，铺满视口，模拟 Excel) */
  gridColRange(scrollX: number, viewW: number): [number, number] {
    return [Math.max(0, this.colAt(scrollX)), this.colAt(scrollX + viewW) + 1]
  }
  gridRowRange(scrollY: number, viewH: number): [number, number] {
    return [Math.max(0, this.rowAt(scrollY)), this.rowAt(scrollY + viewH) + 1]
  }
}

/** 在升序前缀和数组中找最大的 i 使 arr[i] <= value */
function clampSearch(arr: number[], value: number, count: number): number {
  if (value <= 0) return 0
  if (value >= arr[count]) return count - 1
  let lo = 0
  let hi = count
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (arr[mid] <= value) lo = mid + 1
    else hi = mid
  }
  return Math.max(0, lo - 1)
}

/** 0-based 列号 → Excel 列字母(A, B, ..., Z, AA, ...) */
export function colIndexToLetters(index: number): string {
  let n = index + 1
  let s = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}
