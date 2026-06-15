/**
 * Canvas 主渲染器。组合几何/合并/冻结/条件格式/格式引擎，按可视区重绘。
 * 渲染顺序(每个 pane 内): 网格线 → 填充/条件背景 → 数据条 → 边框 → 文本/图标 → 筛选按钮。
 * 表头(行号/列字母)最后绘制，覆盖在最上层。
 */
import type { BorderEdge, CellModel, CellStyle, CellStyleFn, CellStyleOverride, MergeRange, SheetModel, Sparkline, WorkbookModel } from '../model/types'
import { cellKey } from '../model/types'
import { cellDisplayText } from '../model/data-access'
import { mergeStyleOverride } from '../model/mutations'
import { type ViewerTheme, mergeTheme } from './theme'
import { GridMetrics, colIndexToLetters } from '../layout/grid-metrics'
import { MergeIndex } from '../layout/merges'
import { computeFreeze, type FreezeGeometry } from '../layout/freeze'
import { computeViewport, type Pane } from '../layout/viewport'
import { ConditionalEngine, type CellEffect } from './conditional'
import { formatValue } from '../format/number-format'
import { paintFill } from './fills'
import { drawEdge, drawDiagonalEdge, heavierEdge } from './borders'
import {
  fontToCss, measureWidth, resolveHAlign, shrinkScale, wrapLines,
  LINE_HEIGHT_FACTOR, CELL_PADDING,
} from './text'
import { autoFitRowHeights } from '../layout/autofit'
import { drawFilterButton, filterButtonBox, isFilterHeader } from './autofilter'
import { drawPivotToggle, pivotToggleBox } from './pivot-toggle'

/**
 * WPS 单元格内嵌图(DISPIMG)在格内的贴合方式:
 * - `contain`(默认,与 WPS 渲染一致):等比缩放完整显示,周围留白(不裁剪、不变形)。
 *   WPS 打开导出文件时 DISPIMG 固定按 contain 渲染,故默认 contain 让预览所见即所得。
 * - `fill`:拉伸铺满整格,随宽高变形(预览铺满,但导出后 WPS 仍按 contain 显示,二者不一致);
 * - `cover`:等比放大铺满,超出部分裁掉(不变形、不留白)。
 */
export type CellImageFit = 'fill' | 'contain' | 'cover'

export interface RendererOptions {
  theme?: Partial<ViewerTheme>
  cellStyle?: CellStyleFn
  /** 异步内容(WPS 单元格内嵌图)解码完成后请求重绘;壳/控制器接到 scheduleRender */
  onNeedsRedraw?: () => void
  /** WPS 单元格内嵌图贴合方式(默认 contain,与 WPS 渲染一致) */
  cellImageFit?: CellImageFit
  /**
   * 只读单元格视觉钩子 (Phase C, 2026-06-08):
   *   - false (默认) = 无视觉差异 (老行为不变)
   *   - true = 套内置默认 (灰底 #f5f7fa)
   *   - CellStyleOverride 对象 = 固定样式给所有只读格
   *   - CellStyleFn 函数 = 按格自定义
   * 仅在 cellStyle 钩子之后, 该格 editable=false 时套用.
   */
  readOnlyCellStyle?: boolean | CellStyleOverride | CellStyleFn
  /**
   * 查询某格是否可编辑 (Phase C, 2026-06-08):
   * 让渲染器把 ctx.editable 喂给 cellStyle 钩子 + 决定是否套 readOnlyCellStyle.
   * controller 注入, 默认 () => true (不知道权限 → 当全可编辑, 老行为不变).
   */
  isEditable?: (row: number, col: number) => boolean
}

/** 导出为离屏 canvas 的选项 */
export interface ExportToCanvasOptions {
  /** 0-based 闭区间;缺省 = 整个已用范围 */
  range?: MergeRange
  /** 设备像素缩放(越大越清晰),默认 2;受 maxPixels 限制可能被下调 */
  scale?: number
  /** 是否含行号/列字母表头,默认 false(同 Excel 打印默认) */
  includeHeaders?: boolean
  /** 覆盖网格线显隐(缺省跟随 sheet.showGridLines) */
  gridlines?: boolean
  /** 背景色,默认白 */
  background?: string
  /** 单维度设备像素安全上限,默认 16384(超大表自动降 scale) */
  maxPixels?: number
}

/** exportToCanvas 的产物 + 合成叠加层(图片/图表/形状)所需的坐标信息 */
export interface ExportToCanvasResult {
  canvas: HTMLCanvasElement
  /** 实际用的设备像素缩放(可能被 maxPixels 下调) */
  scale: number
  /** 解析出的导出范围 */
  range: MergeRange
  /** 正文宽高(css px,不含表头,zoom=1) */
  bodyW: number
  bodyH: number
  /** 导出空间几何(zoom=1),供合成叠加层定位 */
  metrics: GridMetrics
  /** 画布内"范围左上角格子"所在的 css 像素偏移(含表头时为表头宽高,否则 0) */
  originX: number
  originY: number
  /** 范围左上角在网格坐标系的像素(zoom=1);合成时: deviceX = (originX + gridLeft - gridOriginX) * scale */
  gridOriginX: number
  gridOriginY: number
}

export interface ViewState {
  scrollX: number
  scrollY: number
  width: number // css px
  height: number
  zoom: number
}

/** 矢量导出: 一个单元格解析后的绘制信息 */
export interface CellDrawInfo {
  style: CellStyle
  text: string
  /** 有效字色(已并入条件格式/数字格式/超链接色) */
  color: string
  bold: boolean
  /** 条件格式可矢量化的效果(背景色/数据条/图标);无则省略 */
  effect?: {
    fillColor?: string
    dataBar?: { ratio: number; color: string; gradient: boolean }
    icon?: { setName: string; level: number; count: number }
  }
  /** 矢量层仍画不动的效果(迷你图/旋转/富文本)→ 调用方栅格兜底 */
  complex: boolean
}

export class CanvasRenderer {
  metrics: GridMetrics
  private merges: MergeIndex
  private freeze: FreezeGeometry
  private cond: ConditionalEngine
  private ctx: CanvasRenderingContext2D
  private dpr = 1
  private sparklineIndex = new Map<string, Sparkline>()
  private theme: ViewerTheme
  private cellStyleHook?: CellStyleFn
  private onNeedsRedraw?: () => void
  private cellImageFit: CellImageFit
  /** Phase C 2026-06-08: 只读视觉钩子, 渲染时按格套用 */
  private readOnlyStyleHook?: boolean | CellStyleOverride | CellStyleFn
  /** Phase C 2026-06-08: 查询该格是否可编辑 (controller 注入). 默认全可编辑 (老行为) */
  private isEditableFn: (row: number, col: number) => boolean = () => true
  /** 虚拟外推行/列数(滚动出空行用;0 = 仅按 dimension)。透传给 GridMetrics,不影响导出。 */
  private virtualRows = 0
  private virtualCols = 0
  /** WPS 单元格内嵌图解码缓存: blob src → 已加载的 HTMLImageElement(complete 才画) */
  private cellImageCache = new Map<string, HTMLImageElement>()

  constructor(
    private canvas: HTMLCanvasElement,
    private sheet: SheetModel,
    private workbook: WorkbookModel,
    zoom = 1,
    opts?: RendererOptions,
  ) {
    this.theme = mergeTheme(opts?.theme)
    this.cellStyleHook = opts?.cellStyle
    this.onNeedsRedraw = opts?.onNeedsRedraw
    this.cellImageFit = opts?.cellImageFit ?? 'contain' // 默认 contain:与 WPS DISPIMG 渲染一致
    this.readOnlyStyleHook = opts?.readOnlyCellStyle
    if (opts?.isEditable) this.isEditableFn = opts.isEditable
    const ctx = canvas.getContext('2d')
    if (!ctx) throw new Error('无法获取 canvas 2d context')
    this.ctx = ctx
    // 先做自动行高(撑高换行溢出的行)，再据此建几何
    autoFitRowHeights(sheet, workbook, ctx)
    this.metrics = new GridMetrics(sheet, zoom, this.virtualRows, this.virtualCols)
    this.merges = new MergeIndex(sheet)
    this.freeze = computeFreeze(sheet, this.metrics)
    this.cond = new ConditionalEngine(sheet)
    for (const sp of sheet.sparklines) this.sparklineIndex.set(cellKey(sp.row, sp.col), sp)
  }

  /** 改 WPS 单元格内嵌图贴合方式(fill/contain/cover);返回是否有变化(变了需重绘)。 */
  setCellImageFit(fit: CellImageFit): boolean {
    if (fit === this.cellImageFit) return false
    this.cellImageFit = fit
    return true
  }

  /** 改变缩放: 重建几何(列宽行高表头按 zoom 缩放)。合并/条件格式无需重建。 */
  setZoom(zoom: number): void {
    if (zoom === this.metrics.zoom) return
    this.metrics = new GridMetrics(this.sheet, zoom, this.virtualRows, this.virtualCols)
    this.freeze = computeFreeze(this.sheet, this.metrics)
  }

  /**
   * 设虚拟外推行/列数(滚动出空行/空列);仅当变化时重建 GridMetrics,返回是否变了(变了需重绘 + 刷 spacer)。
   * 不影响 dimension / 导出 / data-access。
   */
  setVirtualExtent(rows: number, cols: number): boolean {
    if (rows === this.virtualRows && cols === this.virtualCols) return false
    this.virtualRows = rows
    this.virtualCols = cols
    this.metrics = new GridMetrics(this.sheet, this.metrics.zoom, rows, cols)
    this.freeze = computeFreeze(this.sheet, this.metrics)
    return true
  }

  /** 内容总尺寸(含表头)，给外层滚动容器用(用虚拟范围 → 可滚动出空行/列) */
  get contentWidth(): number {
    return this.metrics.rowHeaderWidth + this.metrics.virtualWidth
  }
  get contentHeight(): number {
    return this.metrics.colHeaderHeight + this.metrics.virtualHeight
  }
  get freezeGeometry(): FreezeGeometry {
    return this.freeze
  }
  /** 默认行高(已含缩放),供翻页步长估算 */
  get defaultRowPx(): number {
    return this.sheet.defaultRowHeight * this.metrics.zoom
  }

  /** 把行列号映射到当前视图的屏幕矩形(供叠加层定位) */
  screenRectOfCell(view: ViewState, row: number, col: number): { x: number; y: number; w: number; h: number } {
    const x = this.metrics.rowHeaderWidth + this.metrics.colLeft(col) - view.scrollX
    const y = this.metrics.colHeaderHeight + this.metrics.rowTop(row) - view.scrollY
    return { x, y, w: this.metrics.colWidth(col), h: this.metrics.rowHeight(row) }
  }

  // ---------------- 交互查询 ----------------
  private selection: MergeRange | null = null
  setSelection(sel: MergeRange | null): void {
    this.selection = sel
  }
  /** 自动填充柄可见(= editable;1.10.0)。控制器创建后设置。 */
  showFillHandle = false
  /** 自动填充拖拽预览区(目标范围;拖拽中由控制器设)。 */
  private fillPreview: MergeRange | null = null
  setFillPreview(range: MergeRange | null): void {
    this.fillPreview = range
  }
  /** 不连续多选的附加区域(1.13.0;不含 active 区,active 区仍走 setSelection)。 */
  private extraSelection: MergeRange[] = []
  setExtraSelection(ranges: MergeRange[]): void {
    this.extraSelection = ranges
  }

  // ---------------- 查找 ----------------
  private findHits: { row: number; col: number }[] = []
  private findCurrent = -1
  /** 扫描非空单元格,返回命中的格(按阅读顺序: 先行后列) */
  searchCells(query: string, opts: { matchCase?: boolean; wholeCell?: boolean } = {}): { row: number; col: number }[] {
    if (!query) return []
    const q = opts.matchCase ? query : query.toLowerCase()
    const out: { row: number; col: number }[] = []
    for (const cell of this.sheet.cells.values()) {
      if (cell.type === 'empty') continue
      let text = this.cellText(cell.row, cell.col)
      if (!text) continue
      if (!opts.matchCase) text = text.toLowerCase()
      if (opts.wholeCell ? text === q : text.includes(q)) out.push({ row: cell.row, col: cell.col })
    }
    out.sort((a, b) => a.row - b.row || a.col - b.col)
    return out
  }
  /** 设置高亮命中集 + 当前项(供 drawFind 绘制) */
  setFind(hits: { row: number; col: number }[], current: number): void {
    this.findHits = hits
    this.findCurrent = current
  }

  // ---------------- 筛选 ----------------
  private filteredCols = new Set<number>()
  /** 标记哪些列处于筛选态(下拉按钮画成蓝色漏斗) */
  setFilteredCols(cols: Set<number>): void {
    this.filteredCols = cols
  }

  /** 单元格的当前屏幕矩形(冻结感知: 冻结行/列不随滚动) */
  cellScreenRect(view: ViewState, row: number, col: number): { x: number; y: number; w: number; h: number } {
    const hw = this.metrics.rowHeaderWidth
    const hh = this.metrics.colHeaderHeight
    const x = hw + this.metrics.colLeft(col) - (col < this.freeze.frozenCols ? 0 : view.scrollX)
    const y = hh + this.metrics.rowTop(row) - (row < this.freeze.frozenRows ? 0 : view.scrollY)
    return { x, y, w: this.metrics.colWidth(col), h: this.metrics.rowHeight(row) }
  }

  /** 屏幕坐标是否落在某个自动筛选表头的下拉按钮上;是则返回列号 */
  filterButtonAt(view: ViewState, px: number, py: number): number | null {
    const af = this.sheet.autoFilterRange
    if (!af) return null
    const cell = this.cellAtScreen(view, px, py)
    if (!cell || cell.row !== af.top || cell.col < af.left || cell.col > af.right) return null
    const rect = this.cellScreenRect(view, af.top, cell.col)
    const box = filterButtonBox(rect.x, rect.y, rect.w, rect.h)
    if (!box) return null
    if (px >= box.x && px <= box.x + box.size && py >= box.y && py <= box.y + box.size) return cell.col
    return null
  }

  /** 点击是否落在"活动格的数据验证下拉箭头"上(箭头只画在选区左上的列表验证格);是则返回该格。 */
  dataValidationButtonAt(view: ViewState, px: number, py: number): { row: number; col: number } | null {
    const sel = this.selection
    if (!sel) return null
    const row = sel.top, col = sel.left
    if (!this.inDataValidation(row, col)) return null
    const cell = this.cellAtScreen(view, px, py)
    if (!cell || cell.row !== row || cell.col !== col) return null
    const rect = this.cellScreenRect(view, row, col)
    const box = filterButtonBox(rect.x, rect.y, rect.w, rect.h)
    if (!box) return null
    if (px >= box.x && px <= box.x + box.size && py >= box.y && py <= box.y + box.size) return { row, col }
    return null
  }

  /** 屏幕坐标 → 单元格(0-based)。落在表头或越界返回 null。 */
  cellAtScreen(view: ViewState, x: number, y: number): { row: number; col: number } | null {
    const hw = this.metrics.rowHeaderWidth
    const hh = this.metrics.colHeaderHeight
    if (x < hw || y < hh) return null
    const fw = this.freeze.frozenWidth
    const fh = this.freeze.frozenHeight
    const cx = x < hw + fw ? x - hw : x - hw + view.scrollX
    const cy = y < hh + fh ? y - hh : y - hh + view.scrollY
    // 夹到虚拟范围(vCols/vRows-1) → 允许点选滚动出来的空行/空列
    const col = Math.min(Math.max(this.metrics.colAt(cx), 0), this.metrics.vCols - 1)
    const row = Math.min(Math.max(this.metrics.rowAt(cy), 0), this.metrics.vRows - 1)
    return { row, col }
  }

  /** 屏幕 y → 行(用于点行号选整行);落在列表头返回 -1 */
  rowAtScreen(view: ViewState, y: number): number {
    const hh = this.metrics.colHeaderHeight
    if (y < hh) return -1
    const fh = this.freeze.frozenHeight
    const cy = y < hh + fh ? y - hh : y - hh + view.scrollY
    return Math.min(Math.max(this.metrics.rowAt(cy), 0), this.metrics.vRows - 1)
  }
  /** 屏幕 x → 列(用于点列标选整列);落在行表头返回 -1 */
  colAtScreen(view: ViewState, x: number): number {
    const hw = this.metrics.rowHeaderWidth
    if (x < hw) return -1
    const fw = this.freeze.frozenWidth
    const cx = x < hw + fw ? x - hw : x - hw + view.scrollX
    return Math.min(Math.max(this.metrics.colAt(cx), 0), this.metrics.vCols - 1)
  }

  /** 该 cell 所属的合并区(用于选区/复制时按整块处理) */
  mergeAt(row: number, col: number): MergeRange | null {
    return this.merges.rangeOf(row, col) ?? null
  }

  /** 双击列边界: 把列宽自适应到该列内容最宽处 */
  autoFitColumn(col: number): void {
    const ctx = this.ctx
    const zoom = this.metrics.zoom
    let maxW = 0
    for (const cell of this.sheet.cells.values()) {
      if (cell.col !== col || cell.type === 'empty') continue
      if (this.merges.rangeOf(cell.row, col)) continue
      const text = this.cellText(cell.row, col)
      if (!text) continue
      ctx.font = fontToCss(this.styleOf(cell).font, zoom)
      const w = ctx.measureText(text).width
      if (w > maxW) maxW = w
    }
    const widthUnzoomed =
      maxW === 0 ? this.sheet.defaultColWidth : Math.max(16, Math.ceil(maxW / zoom) + CELL_PADDING * 2 + 3)
    const info = this.sheet.columns.get(col)
    this.sheet.columns.set(col, { width: widthUnzoomed, hidden: info?.hidden ?? false })
    this.rebuildMetrics()
  }

  /** 双击行边界: 行高自适应到该行内容最高处 */
  autoFitRow(row: number): void {
    const ctx = this.ctx
    let maxH = this.sheet.defaultRowHeight
    for (const cell of this.sheet.cells.values()) {
      if (cell.row !== row || cell.type === 'empty') continue
      if (this.merges.rangeOf(row, cell.col)) continue
      const style = this.styleOf(cell)
      const text = this.cellText(row, cell.col)
      if (!text) continue
      const lineHpx = style.font.size * (96 / 72) * LINE_HEIGHT_FACTOR
      let lines = 1
      if (style.wrapText) {
        const z = this.metrics.zoom
        const availW = this.metrics.colWidth(cell.col) / z - CELL_PADDING * 2
        ctx.font = fontToCss(style.font, 1)
        lines = wrapLines(ctx, text, fontToCss(style.font, 1), availW).length
      } else {
        lines = text.split('\n').length
      }
      const h = lines * lineHpx + CELL_PADDING * 2
      if (h > maxH) maxH = h
    }
    const info = this.sheet.rows.get(row)
    this.sheet.rows.set(row, { height: Math.ceil(maxH), hidden: info?.hidden ?? false })
    this.rebuildMetrics()
  }

  /** 拖拽改列宽(传入屏幕像素,内部换算为非缩放存储) */
  setColWidthPx(col: number, px: number): void {
    const info = this.sheet.columns.get(col)
    this.sheet.columns.set(col, { width: Math.max(8, px / this.metrics.zoom), hidden: info?.hidden ?? false })
    this.rebuildMetrics()
  }
  /** 拖拽改行高 */
  setRowHeightPx(row: number, px: number): void {
    const info = this.sheet.rows.get(row)
    this.sheet.rows.set(row, { height: Math.max(6, px / this.metrics.zoom), hidden: info?.hidden ?? false })
    this.rebuildMetrics()
  }

  /** 复制带格式用: 单元格的内联 CSS(粗体/斜体/色/底色/对齐) */
  cellInlineStyle(row: number, col: number): string {
    const cell = this.sheet.cells.get(cellKey(row, col))
    if (!cell) return ''
    const s = this.styleOf(cell)
    const css: string[] = []
    if (s.font.bold) css.push('font-weight:bold')
    if (s.font.italic) css.push('font-style:italic')
    if (s.font.color && s.font.color !== '#000000') css.push('color:' + s.font.color)
    if (s.fill.type === 'solid' && s.fill.fgColor) css.push('background:' + s.fill.fgColor)
    if (s.hAlign === 'center' || s.hAlign === 'right') css.push('text-align:' + s.hAlign)
    return css.join(';')
  }

  /** 列宽/行高变化后重建几何 */
  rebuildMetrics(): void {
    this.metrics = new GridMetrics(this.sheet, this.metrics.zoom, this.virtualRows, this.virtualCols)
    this.freeze = computeFreeze(this.sheet, this.metrics)
    this.merges = new MergeIndex(this.sheet) // 结构编辑(增删行列)会改 merges → 同步重建索引
  }

  /** 选区统计: 遍历非空 cell(O(非空数),不随选区大小爆炸) */
  selectionStats(range: MergeRange): { count: number; numCount: number; sum: number; avg: number; min: number; max: number } {
    let count = 0
    let numCount = 0
    let sum = 0
    let min = Infinity
    let max = -Infinity
    for (const cell of this.sheet.cells.values()) {
      if (cell.row < range.top || cell.row > range.bottom || cell.col < range.left || cell.col > range.right) continue
      if (cell.type === 'empty') continue
      if (this.cellText(cell.row, cell.col) !== '') count++
      if (typeof cell.raw === 'number') {
        numCount++
        sum += cell.raw
        if (cell.raw < min) min = cell.raw
        if (cell.raw > max) max = cell.raw
      }
    }
    return { count, numCount, sum, avg: numCount ? sum / numCount : 0, min, max }
  }

  /**
   * 矢量导出用: 解析一个单元格的绘制信息(样式 + 显示文本 + 有效字色/粗体 + 是否"复杂")。
   * complex = 含条件格式背景/数据条/图标、迷你图、旋转、富文本 —— 矢量层画不动,调用方应栅格兜底。
   * 返回 null 表示该格无任何可绘制内容(无填充/边框/文本)。
   */
  exportCellDraw(row: number, col: number): CellDrawInfo | null {
    const cell = this.sheet.cells.get(cellKey(row, col))
    if (!cell) return null
    const style = this.styleOf(cell)
    const text = this.cellText(row, col)
    const hasBox = style.fill.type !== 'none' || style.borders.top || style.borders.bottom || style.borders.left || style.borders.right
    if (!text && !hasBox) return null

    const isRich = cell.type === 'richtext'
    const rotation = style.textRotation
    const sparkline = this.sparklineIndex.has(cellKey(row, col))
    const effect = this.cond.hasRules() ? this.cond.effectsFor(row, col, cell.raw ?? null) : null
    // 条件背景/数据条/图标可矢量画;迷你图/旋转/富文本仍需栅格兜底
    const complex = isRich || (!!rotation && rotation !== 0) || sparkline

    let color = style.font.color
    if (!isRich) {
      const formatted = formatValue(cell.raw, style.numFmt, this.workbook.date1904)
      color =
        effect?.fontColor ||
        formatted.color ||
        (cell.type === 'hyperlink' ? this.workbook.themeColors[10] || '#0563C1' : style.font.color)
    }
    const vecEffect =
      effect && (effect.fillColor || effect.dataBar || effect.icon)
        ? { fillColor: effect.fillColor, dataBar: effect.dataBar, icon: effect.icon }
        : undefined
    return { style, text, color, bold: style.font.bold || !!effect?.bold, effect: vecEffect, complex }
  }

  /** 单元格的显示文本(套数字格式后);空返回 '' */
  cellText(row: number, col: number): string {
    const cell = this.sheet.cells.get(cellKey(row, col))
    // 用 styleOf(含 cellStyle 钩子)的 numFmt;复用纯函数 cellDisplayText 保持单一真相源
    return cellDisplayText(cell, cell ? this.styleOf(cell) : undefined, this.workbook.date1904)
  }

  /** 单元格的公式文本(无则 null),供公式栏显示 */
  cellFormula(row: number, col: number): string | null {
    const cell = this.sheet.cells.get(cellKey(row, col))
    return cell?.formula ? '=' + cell.formula : null
  }

  /** 单元格原始数值(供状态栏统计);非数字返回 null */
  cellNumber(row: number, col: number): number | null {
    const cell = this.sheet.cells.get(cellKey(row, col))
    return cell && typeof cell.raw === 'number' ? cell.raw : null
  }

  cellHyperlink(row: number, col: number): string | null {
    const cell = this.sheet.cells.get(cellKey(row, col))
    return cell?.hyperlink ?? null
  }

  /** 单元格批注(无则 null) */
  commentAt(row: number, col: number): string | null {
    const cell = this.sheet.cells.get(cellKey(row, col))
    return cell?.comment ?? null
  }

  private inDataValidation(row: number, col: number): boolean {
    return this.sheet.dataValidations.some(
      (rg) => row >= rg.top && row <= rg.bottom && col >= rg.left && col <= rg.right,
    )
  }

  /** 若该 cell 文本被裁切(非换行、宽度不足、且溢出也放不下),返回完整文本供 tooltip;否则 null */
  overflowTextAt(row: number, col: number): string | null {
    if (this.merges.rangeOf(row, col)) return null
    const cell = this.sheet.cells.get(cellKey(row, col))
    if (!cell || cell.type === 'empty') return null
    const style = this.styleOf(cell)
    if (style.wrapText) return null
    const text = this.cellText(row, col)
    if (!text) return null
    this.ctx.font = fontToCss(style.font, this.metrics.zoom)
    const tw = this.ctx.measureText(text).width
    const pad = CELL_PADDING * 2 * this.metrics.zoom
    const w = this.metrics.colWidth(col)
    if (tw <= w - pad) return null
    // 文本可能溢出到相邻空格 → 若溢出后能放下，就不需要 tooltip
    const isNumber = cell.type === 'number' || (cell.type === 'formula' && typeof cell.raw === 'number')
    const spill = this.spillClip(row, col, 0, w, tw - (w - pad), resolveHAlign(style.hAlign, isNumber))
    return tw > spill.w - pad + 1 ? text : null
  }

  render(view: ViewState): void {
    this.dpr = window.devicePixelRatio || 1
    // canvas 是"替换元素":CSS 的 inset:0 / width:100% 对它无效 —— width:auto 会解析成它的固有尺寸
    // (= 缓冲像素数 width*dpr)。dpr≠1(系统缩放 125%/150% / 浏览器 Ctrl+缩放)时,canvas 会以
    // width*dpr 个 CSS 像素显示,比容器大 dpr 倍 → 整个网格被放大,和 DOM 叠加层(浮动图/图表)及
    // 鼠标命中错位,且越往右下偏得越多。必须显式把 CSS 显示尺寸钉成逻辑尺寸(width/height),
    // 缓冲(width*dpr)再被浏览器降采样显示 → 高清且与逻辑坐标 1:1 对齐。这是 HiDPI canvas 的标准做法。
    const cw = view.width + 'px'
    const ch = view.height + 'px'
    if (this.canvas.style.width !== cw) this.canvas.style.width = cw
    if (this.canvas.style.height !== ch) this.canvas.style.height = ch
    this.paint(view, { headers: true, pageBreaks: true, selection: true })
  }

  /**
   * 绘制核心: 被实时 render() 与导出 exportToCanvas() 共用。
   * 用 this.{canvas,ctx,metrics,freeze,selection,dpr} 当前值,flags 控制 UI 装饰是否绘制。
   */
  private paint(
    view: ViewState,
    flags: { headers: boolean; pageBreaks: boolean; selection: boolean },
    background = '#FFFFFF',
  ): void {
    const { width, height, zoom } = view
    // 调整 canvas 像素尺寸(高清/导出 scale)
    const pw = Math.round(width * this.dpr)
    const ph = Math.round(height * this.dpr)
    if (this.canvas.width !== pw || this.canvas.height !== ph) {
      this.canvas.width = pw
      this.canvas.height = ph
    }
    const ctx = this.ctx
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0)
    ctx.clearRect(0, 0, width, height)
    ctx.fillStyle = background
    ctx.fillRect(0, 0, width, height)

    const layout = computeViewport(this.metrics, this.freeze, view.scrollX, view.scrollY, width, height)

    for (const pane of layout.panes) {
      this.drawPane(pane, zoom)
    }

    if (flags.headers) this.drawHeaders(layout.panes, view)
    this.drawFreezeLines(layout) // 导出时 freeze=0 → 直接返回
    if (flags.pageBreaks) this.drawPageBreaks(view)
    if (flags.selection) this.drawPivotToggles(view) // 折叠按钮只在实时视图画,导出件不画
    if (flags.selection && this.findHits.length) this.drawFind(view)
    if (flags.selection) this.drawSelection(view)
    if (flags.selection) this.drawFillPreview(view)
  }

  /** 画各透视表的行分组折叠/展开按钮(贴在分组表头格最左)。 */
  private drawPivotToggles(view: ViewState): void {
    const tables = this.sheet.pivotTables
    if (!tables?.length) return
    const hw = this.metrics.rowHeaderWidth
    const hh = this.metrics.colHeaderHeight
    const ctx = this.ctx
    ctx.save()
    ctx.beginPath()
    ctx.rect(hw, hh, view.width - hw, view.height - hh)
    ctx.clip()
    for (const pivot of tables) {
      const groups = pivot.rowGroups
      if (!groups?.length) continue
      const collapsed = new Set(pivot.collapsed ?? [])
      const col = pivot.range.left
      for (const g of groups) {
        const rect = this.cellScreenRect(view, g.row, col)
        if (rect.x + rect.w < hw || rect.y + rect.h < hh || rect.x > view.width || rect.y > view.height) continue
        const box = pivotToggleBox(rect.x, rect.y, rect.w, rect.h)
        if (box) drawPivotToggle(ctx, box.x, box.y, box.size, collapsed.has(g.key))
      }
    }
    ctx.restore()
  }

  /** 屏幕坐标是否落在某透视表的折叠按钮上;是则返回 { tableIdx, key }。 */
  pivotToggleAt(view: ViewState, px: number, py: number): { tableIdx: number; key: string } | null {
    const tables = this.sheet.pivotTables
    if (!tables?.length) return null
    const cell = this.cellAtScreen(view, px, py)
    if (!cell) return null
    for (let ti = 0; ti < tables.length; ti++) {
      const pivot = tables[ti]
      const groups = pivot.rowGroups
      if (!groups?.length || cell.col !== pivot.range.left) continue
      for (const g of groups) {
        if (g.row !== cell.row) continue
        const rect = this.cellScreenRect(view, g.row, pivot.range.left)
        const box = pivotToggleBox(rect.x, rect.y, rect.w, rect.h)
        if (box && px >= box.x && px <= box.x + box.size && py >= box.y && py <= box.y + box.size) return { tableIdx: ti, key: g.key }
      }
    }
    return null
  }

  /** 查找高亮: 所有命中淡黄,当前项橙色描边 */
  private drawFind(view: ViewState): void {
    const ctx = this.ctx
    const hw = this.metrics.rowHeaderWidth
    const hh = this.metrics.colHeaderHeight
    ctx.save()
    ctx.beginPath()
    ctx.rect(hw, hh, view.width - hw, view.height - hh)
    ctx.clip()
    for (let i = 0; i < this.findHits.length; i++) {
      const hit = this.findHits[i]
      const mg = this.merges.rangeOf(hit.row, hit.col)
      const top = mg ? mg.top : hit.row
      const left = mg ? mg.left : hit.col
      const tl = this.screenRectOfCell(view, top, left)
      const w = mg ? this.metrics.colLeft(mg.right + 1) - this.metrics.colLeft(mg.left) : tl.w
      const h = mg ? this.metrics.rowTop(mg.bottom + 1) - this.metrics.rowTop(mg.top) : tl.h
      if (tl.x + w < hw || tl.y + h < hh || tl.x > view.width || tl.y > view.height) continue
      ctx.fillStyle = i === this.findCurrent ? 'rgba(255,145,0,0.40)' : 'rgba(255,231,76,0.50)'
      ctx.fillRect(tl.x, tl.y, w, h)
      if (i === this.findCurrent) {
        ctx.strokeStyle = '#F57C00'
        ctx.lineWidth = 2
        ctx.strokeRect(Math.round(tl.x) + 1, Math.round(tl.y) + 1, Math.round(w) - 2, Math.round(h) - 2)
      }
    }
    ctx.restore()
  }

  /**
   * 把整表(或指定区域)渲染到一张离屏 canvas: 无选区高亮/分页线/冻结分区,scale 可控。
   * 复用全部单元格绘制逻辑(填充/边框/条件格式/迷你图/文本)。图片/图表/形状是 DOM 叠加层,
   * 不在此绘制 —— 由调用方拿到 result.metrics/scale/origin 后合成到底图上。
   */
  exportToCanvas(opts: ExportToCanvasOptions = {}): ExportToCanvasResult {
    const dim = this.sheet.dimension
    const range = opts.range ?? {
      top: 0,
      left: 0,
      bottom: Math.max(0, dim.rows - 1),
      right: Math.max(0, dim.cols - 1),
    }
    const includeHeaders = opts.includeHeaders ?? false
    const background = opts.background ?? '#FFFFFF'

    // 导出几何固定 zoom=1(自然尺寸,与屏幕缩放无关);清晰度靠 scale
    const m = new GridMetrics(this.sheet, 1)
    const hw = m.rowHeaderWidth
    const hh = m.colHeaderHeight
    const gridOriginX = m.colLeft(range.left)
    const gridOriginY = m.rowTop(range.top)
    const bodyW = m.colLeft(range.right + 1) - gridOriginX
    const bodyH = m.rowTop(range.bottom + 1) - gridOriginY

    // 画布(含表头带)css 尺寸 → 据此 + scale 定设备像素;裁掉表头时事后 crop
    const cssW = hw + bodyW
    const cssH = hh + bodyH
    const cap = opts.maxPixels ?? 16384
    let scale = opts.scale ?? 2
    scale = Math.min(scale, cap / Math.max(1, cssW), cap / Math.max(1, cssH))
    if (!(scale > 0)) scale = 1

    const tmp = document.createElement('canvas')

    // 临时换实例字段,复用 paint();完事还原
    const saved = {
      canvas: this.canvas,
      ctx: this.ctx,
      metrics: this.metrics,
      freeze: this.freeze,
      selection: this.selection,
      dpr: this.dpr,
      showGridLines: this.sheet.showGridLines,
    }
    this.canvas = tmp
    this.ctx = tmp.getContext('2d')!
    this.metrics = m
    this.freeze = { frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0 }
    this.selection = null
    this.dpr = scale
    if (opts.gridlines !== undefined) this.sheet.showGridLines = opts.gridlines

    try {
      this.paint(
        { scrollX: gridOriginX, scrollY: gridOriginY, width: cssW, height: cssH, zoom: 1 },
        { headers: includeHeaders, pageBreaks: false, selection: false },
        background,
      )
    } finally {
      this.canvas = saved.canvas
      this.ctx = saved.ctx
      this.metrics = saved.metrics
      this.freeze = saved.freeze
      this.selection = saved.selection
      this.dpr = saved.dpr
      this.sheet.showGridLines = saved.showGridLines
    }

    if (includeHeaders) {
      return { canvas: tmp, scale, range, bodyW, bodyH, metrics: m, originX: hw, originY: hh, gridOriginX, gridOriginY }
    }
    // 无表头: 裁掉左上的表头带
    const out = document.createElement('canvas')
    out.width = Math.max(1, Math.round(bodyW * scale))
    out.height = Math.max(1, Math.round(bodyH * scale))
    const octx = out.getContext('2d')!
    octx.fillStyle = background
    octx.fillRect(0, 0, out.width, out.height)
    octx.drawImage(tmp, Math.round(hw * scale), Math.round(hh * scale), out.width, out.height, 0, 0, out.width, out.height)
    return { canvas: out, scale, range, bodyW, bodyH, metrics: m, originX: 0, originY: 0, gridOriginX, gridOriginY }
  }

  /** 手动分页符: 蓝色虚线(在正文区内,随内容滚动) */
  private drawPageBreaks(view: ViewState): void {
    const pb = this.sheet.pageBreaks
    if (!pb || (!pb.rows.length && !pb.cols.length)) return
    const ctx = this.ctx
    const hw = this.metrics.rowHeaderWidth
    const hh = this.metrics.colHeaderHeight
    ctx.save()
    ctx.beginPath()
    ctx.rect(hw, hh, view.width - hw, view.height - hh)
    ctx.clip()
    ctx.strokeStyle = '#4472C4'
    ctx.lineWidth = 1
    ctx.setLineDash([4, 3])
    for (const c of pb.cols) {
      const sx = Math.round(hw + this.metrics.colLeft(c) - view.scrollX) + 0.5
      if (sx <= hw) continue
      ctx.beginPath()
      ctx.moveTo(sx, hh)
      ctx.lineTo(sx, view.height)
      ctx.stroke()
    }
    for (const r of pb.rows) {
      const sy = Math.round(hh + this.metrics.rowTop(r) - view.scrollY) + 0.5
      if (sy <= hh) continue
      ctx.beginPath()
      ctx.moveTo(hw, sy)
      ctx.lineTo(view.width, sy)
      ctx.stroke()
    }
    ctx.restore()
  }

  /** 选区高亮: 半透明填充 + 蓝色边框，裁到表头以下的正文区。 */
  private drawSelection(view: ViewState): void {
    const sel = this.selection
    if (!sel) return
    const ctx = this.ctx
    const hw = this.metrics.rowHeaderWidth
    const hh = this.metrics.colHeaderHeight
    const tl = this.screenRectOfCell(view, sel.top, sel.left)
    const br = this.screenRectOfCell(view, sel.bottom, sel.right)
    const x = tl.x
    const y = tl.y
    const w = br.x + br.w - tl.x
    const h = br.y + br.h - tl.y

    ctx.save()
    ctx.beginPath()
    ctx.rect(hw, hh, view.width - hw, view.height - hh)
    ctx.clip()
    // 不连续多选的附加区域:先画(填充 + 边框,无填充柄;1.13.0)
    for (const er of this.extraSelection) {
      const etl = this.screenRectOfCell(view, er.top, er.left)
      const ebr = this.screenRectOfCell(view, er.bottom, er.right)
      ctx.fillStyle = this.theme.selFill
      ctx.fillRect(etl.x, etl.y, ebr.x + ebr.w - etl.x, ebr.y + ebr.h - etl.y)
      ctx.strokeStyle = this.theme.selBorder
      ctx.lineWidth = 2
      ctx.strokeRect(Math.round(etl.x) + 1, Math.round(etl.y) + 1, Math.round(ebr.x + ebr.w - etl.x) - 2, Math.round(ebr.y + ebr.h - etl.y) - 2)
    }
    // 单格选区不填充(像 Excel 的活动单元格),多格才铺淡蓝;多选时活动区也铺底好区分
    const single = sel.top === sel.bottom && sel.left === sel.right && this.extraSelection.length === 0
    if (!single) {
      ctx.fillStyle = this.theme.selFill
      ctx.fillRect(x, y, w, h)
    }
    ctx.strokeStyle = this.theme.selBorder
    ctx.lineWidth = 2
    ctx.strokeRect(Math.round(x) + 1, Math.round(y) + 1, Math.round(w) - 2, Math.round(h) - 2)
    // 自动填充柄:选区右下角的小方块(editable 才画;1.10.0)。不连续多选时不画(对齐 Excel)
    if (this.showFillHandle && this.extraSelection.length === 0) {
      const hr = this.fillHandleRect(view)
      if (hr) {
        ctx.fillStyle = this.theme.selBorder
        ctx.strokeStyle = '#fff'
        ctx.lineWidth = 1
        ctx.fillRect(hr.x, hr.y, hr.w, hr.h)
        ctx.strokeRect(hr.x + 0.5, hr.y + 0.5, hr.w - 1, hr.h - 1)
      }
    }
    ctx.restore()
  }

  /** 自动填充柄的屏幕矩形(选区右下角的小方块);选区不可见/无选区返 null。 */
  fillHandleRect(view: ViewState): { x: number; y: number; w: number; h: number } | null {
    const sel = this.selection
    if (!sel) return null
    const br = this.screenRectOfCell(view, sel.bottom, sel.right)
    const hw = this.metrics.rowHeaderWidth
    const hh = this.metrics.colHeaderHeight
    const cx = br.x + br.w
    const cy = br.y + br.h
    if (cx < hw || cy < hh || cx > view.width || cy > view.height) return null // 角不在可视区
    const s = 7
    return { x: Math.round(cx - s / 2 - 1), y: Math.round(cy - s / 2 - 1), w: s, h: s }
  }
  /** 点 (px,py) 是否落在填充柄上(命中区比绘制略大,好点中)。 */
  fillHandleAt(view: ViewState, px: number, py: number): boolean {
    if (!this.showFillHandle) return false
    const hr = this.fillHandleRect(view)
    if (!hr) return false
    const pad = 3
    return px >= hr.x - pad && px <= hr.x + hr.w + pad && py >= hr.y - pad && py <= hr.y + hr.h + pad
  }

  /** 拖拽预览:目标范围的虚线框(超出源选区的部分)。 */
  private drawFillPreview(view: ViewState): void {
    const pv = this.fillPreview
    if (!pv) return
    const ctx = this.ctx
    const hw = this.metrics.rowHeaderWidth
    const hh = this.metrics.colHeaderHeight
    const tl = this.screenRectOfCell(view, pv.top, pv.left)
    const br = this.screenRectOfCell(view, pv.bottom, pv.right)
    ctx.save()
    ctx.beginPath()
    ctx.rect(hw, hh, view.width - hw, view.height - hh)
    ctx.clip()
    ctx.strokeStyle = this.theme.selBorder
    ctx.lineWidth = 1
    ctx.setLineDash([4, 3])
    ctx.strokeRect(Math.round(tl.x) + 0.5, Math.round(tl.y) + 0.5, Math.round(br.x + br.w - tl.x) - 1, Math.round(br.y + br.h - tl.y) - 1)
    ctx.restore()
  }

  // ---------------- pane 绘制 ----------------
  private drawPane(pane: Pane, zoom: number): void {
    const ctx = this.ctx
    if (pane.clipW <= 0 || pane.clipH <= 0) return
    ctx.save()
    ctx.beginPath()
    ctx.rect(pane.clipX, pane.clipY, pane.clipW, pane.clipH)
    ctx.clip()

    const hw = this.metrics.rowHeaderWidth
    const hh = this.metrics.colHeaderHeight
    const sx = (c: number) => hw + this.metrics.colLeft(c) - pane.offsetX
    const sy = (r: number) => hh + this.metrics.rowTop(r) - pane.offsetY

    // 网格线铺满整个 pane 视口(可超出数据范围，模拟 Excel 无限网格)。
    // clip 已限制绘制区，超界的线被裁掉，不影响性能。
    const gridC0 = this.metrics.colAt(pane.clipX - hw + pane.offsetX)
    const gridC1 = this.metrics.colAt(pane.clipX + pane.clipW - hw + pane.offsetX) + 1
    const gridR0 = this.metrics.rowAt(pane.clipY - hh + pane.offsetY)
    const gridR1 = this.metrics.rowAt(pane.clipY + pane.clipH - hh + pane.offsetY) + 1

    // 可视区内的合并区(网格线绘制与合并绘制共用)
    const visibleMerges = this.sheet.merges.filter(
      (m) => !(m.bottom < pane.rowStart || m.top > pane.rowEnd || m.right < pane.colStart || m.left > pane.colEnd),
    )

    // 1. 网格线 —— 合并区内部不画(跟 Excel/WPS 一致:合并后内部网格线消失,只留外边界)
    if (this.sheet.showGridLines) {
      ctx.strokeStyle = this.theme.gridLine
      ctx.lineWidth = 1
      ctx.beginPath()
      const yTop = pane.clipY
      const yBot = pane.clipY + pane.clipH
      for (let c = Math.max(0, gridC0); c <= gridC1; c++) {
        const x = Math.round(sx(c)) + 0.5
        // c 落在某合并区内部(m.left < c <= m.right)→ 该竖线在此合并区的纵向区间不画
        const gaps: Array<[number, number]> = []
        for (const m of visibleMerges) if (m.left < c && c <= m.right) gaps.push([sy(m.top), sy(m.bottom + 1)])
        strokeGapped(ctx, x, yTop, yBot, gaps, true)
      }
      const xL = pane.clipX
      const xR = pane.clipX + pane.clipW
      for (let r = Math.max(0, gridR0); r <= gridR1; r++) {
        const y = Math.round(sy(r)) + 0.5
        // r 落在某合并区内部(m.top < r <= m.bottom)→ 该横线在此合并区的横向区间不画
        const gaps: Array<[number, number]> = []
        for (const m of visibleMerges) if (m.top < r && r <= m.bottom) gaps.push([sx(m.left), sx(m.right + 1)])
        strokeGapped(ctx, y, xL, xR, gaps, false)
      }
      ctx.stroke()
    }

    // 2. 合并区(锚点可能在可视区外，需单独扫描)
    const coveredAnchorsDrawn = new Set<string>()
    for (const m of visibleMerges) {
      const x = sx(m.left)
      const y = sy(m.top)
      const w = this.metrics.colLeft(m.right + 1) - this.metrics.colLeft(m.left)
      const h = this.metrics.rowTop(m.bottom + 1) - this.metrics.rowTop(m.top)
      const cell = this.sheet.cells.get(cellKey(m.top, m.left))
      this.paintCellBox(cell, m.top, m.left, x, y, w, h, zoom, true)
      coveredAnchorsDrawn.add(cellKey(m.top, m.left))
    }

    // 3. 普通单元格
    for (let r = pane.rowStart; r <= pane.rowEnd; r++) {
      for (let c = pane.colStart; c <= pane.colEnd; c++) {
        if (this.merges.rangeOf(r, c)) continue // 合并区已在上面整体绘制
        const cell = this.sheet.cells.get(cellKey(r, c))
        const x = sx(c)
        const y = sy(r)
        const w = this.metrics.colWidth(c)
        const h = this.metrics.rowHeight(r)
        this.paintCellBox(cell, r, c, x, y, w, h, zoom)
      }
    }

    ctx.restore()
  }

  /** 取某格指定边的边框样式(供相邻共享边取较重者);越界/无格返回 undefined */
  private borderEdgeOf(row: number, col: number, side: 'top' | 'bottom' | 'left' | 'right'): BorderEdge | undefined {
    if (row < 0 || col < 0 || row >= this.metrics.rows || col >= this.metrics.cols) return undefined
    const cell = this.sheet.cells.get(cellKey(row, col))
    if (!cell) return undefined
    return this.styleOf(cell).borders[side]
  }

  /** 画单个 cell(或合并区)的: 填充 → 条件背景 → 数据条 → 边框 → 内容 → 图标 → 筛选按钮 */
  private paintCellBox(
    cell: CellModel | undefined,
    row: number,
    col: number,
    x: number,
    y: number,
    w: number,
    h: number,
    zoom: number,
    isMerge = false,
  ): void {
    const ctx = this.ctx
    const style = cell ? this.styleOf(cell) : undefined

    // 填充
    if (style && style.fill.type !== 'none') paintFill(ctx, style.fill, x, y, w, h)

    // 条件格式
    let effect: CellEffect | null = null
    if (this.cond.hasRules()) {
      effect = this.cond.effectsFor(row, col, cell?.raw ?? null)
      if (effect?.fillColor) {
        ctx.fillStyle = effect.fillColor
        ctx.fillRect(x, y, w, h)
      }
      if (effect?.dataBar) {
        const barW = Math.max(0, (w - 4) * effect.dataBar.ratio)
        ctx.save()
        if (effect.dataBar.gradient) {
          const g = ctx.createLinearGradient(x, y, x + barW, y)
          g.addColorStop(0, effect.dataBar.color)
          g.addColorStop(1, withAlpha(effect.dataBar.color, 0.3))
          ctx.fillStyle = g
        } else {
          ctx.fillStyle = effect.dataBar.color
        }
        ctx.fillRect(x + 2, y + 2, barW, h - 4)
        ctx.restore()
      }
    }

    // 边框
    if (style) {
      const b = style.borders
      if (isMerge) {
        // 合并区:画自身四周(相邻优先级对多列/多行边界复杂,合并区一般自带边框,直接画)
        drawEdge(ctx, b.top, x, y, x + w, y)
        drawEdge(ctx, b.bottom, x, y + h, x + w, y + h)
        drawEdge(ctx, b.left, x, y, x, y + h)
        drawEdge(ctx, b.right, x + w, y, x + w, y + h)
      } else {
        // 普通格:共享边与相邻格取较重的一条,绘制顺序无关、与 Excel/WPS 一致
        drawEdge(ctx, heavierEdge(b.top, this.borderEdgeOf(row - 1, col, 'bottom')), x, y, x + w, y)
        drawEdge(ctx, heavierEdge(b.bottom, this.borderEdgeOf(row + 1, col, 'top')), x, y + h, x + w, y + h)
        drawEdge(ctx, heavierEdge(b.left, this.borderEdgeOf(row, col - 1, 'right')), x, y, x, y + h)
        drawEdge(ctx, heavierEdge(b.right, this.borderEdgeOf(row, col + 1, 'left')), x + w, y, x + w, y + h)
      }
      // 对角线(↘ 左上→右下 / ↗ 左下→右上),合并区跨整块
      if (b.diagonal && (b.diagonalDown || b.diagonalUp)) {
        if (b.diagonalDown) drawDiagonalEdge(ctx, b.diagonal, x, y, x + w, y + h)
        if (b.diagonalUp) drawDiagonalEdge(ctx, b.diagonal, x, y + h, x + w, y)
      }
    }

    // 图标(条件格式 iconSet)
    let contentX = x
    let contentW = w
    if (effect?.icon) {
      drawIcon(ctx, effect.icon, x + 2, y, h)
      contentX += 18
      contentW -= 18
    }

    // 迷你图(通常在空单元格里)
    const sparkline = this.sparklineIndex.get(cellKey(row, col))
    if (sparkline) drawSparkline(ctx, sparkline, x, y, w, h)

    // 内容: WPS 单元格内嵌图(DISPIMG)优先画图(占文本位置),否则画文本/数值
    if (cell?.dispImgId) {
      this.drawCellImage(cell.dispImgId, x, y, w, h, zoom)
    } else if (cell && cell.type !== 'empty' && style) {
      this.drawCellContent(cell, style, row, col, contentX, y, contentW, h, zoom, effect)
    }

    // 自动筛选下拉
    if (isFilterHeader(this.sheet.autoFilterRange, row, col)) {
      drawFilterButton(ctx, x, y, w, h, this.filteredCols.has(col))
    }

    // 数据验证下拉: 仅在"活动单元格"(选区左上)且属列表验证时显示(同 Excel)
    if (
      this.selection &&
      row === this.selection.top &&
      col === this.selection.left &&
      this.inDataValidation(row, col)
    ) {
      drawFilterButton(ctx, x, y, w, h)
    }

    // 批注标记: 右上角红三角
    if (cell?.comment) {
      const s = 7
      ctx.fillStyle = '#E8453C'
      ctx.beginPath()
      ctx.moveTo(x + w - s, y)
      ctx.lineTo(x + w, y)
      ctx.lineTo(x + w, y + s)
      ctx.closePath()
      ctx.fill()
    }
  }

  /**
   * 把 WPS 单元格内嵌图(DISPIMG)画进格内,贴合方式由 `cellImageFit` 决定(fill/contain/cover);
   * 始终裁剪到格内、随行高列宽变化。未解码完成/缺图时画淡占位。
   */
  private drawCellImage(id: string, x: number, y: number, w: number, h: number, zoom: number): void {
    const ctx = this.ctx
    // 留 ~1px 内边距(不盖网格线);格子太小就不留,优先铺满
    const pad = w > 6 && h > 6 ? Math.min(1 * zoom, 1.5) : 0
    const boxX = x + pad
    const boxY = y + pad
    const boxW = Math.max(1, w - pad * 2)
    const boxH = Math.max(1, h - pad * 2)
    const entry = this.workbook.cellImages?.get(id)
    const el = entry?.src ? this.getCellImageEl(entry.src) : null // 触发异步解码
    if (!el || !el.complete || el.naturalWidth === 0) {
      // 加载中 / 缺图:**不画灰底**(露出单元格自身填充色,通常是白)。缺图时只画个淡图标提示,不盖底色。
      if (!entry?.src) this.drawImagePlaceholder(boxX, boxY, boxW, boxH, zoom)
      return
    }
    // 按 fit 模式算目标矩形:fill=铺满变形;contain=等比留白;cover=等比裁剪铺满
    let dx = boxX
    let dy = boxY
    let dw = boxW
    let dh = boxH
    if (this.cellImageFit !== 'fill') {
      const iw = el.naturalWidth
      const ih = el.naturalHeight
      const scale =
        this.cellImageFit === 'cover' ? Math.max(boxW / iw, boxH / ih) : Math.min(boxW / iw, boxH / ih)
      dw = iw * scale
      dh = ih * scale
      dx = boxX + (boxW - dw) / 2
      dy = boxY + (boxH - dh) / 2
    }
    ctx.save()
    ctx.beginPath()
    ctx.rect(boxX, boxY, boxW, boxH) // 裁剪到格内(cover 超出部分被裁;防图溢出相邻格)
    ctx.clip()
    try {
      ctx.drawImage(el, dx, dy, dw, dh)
    } catch {
      /* 解码失败忽略 */
    }
    ctx.restore()
  }

  /** 取/起一张内嵌图的解码;未缓存则起加载,onload 请求重绘(无 DOM 环境返 null) */
  private getCellImageEl(src: string): HTMLImageElement | null {
    const cached = this.cellImageCache.get(src)
    if (cached) return cached
    if (typeof Image === 'undefined') return null
    const el = new Image()
    this.cellImageCache.set(src, el)
    el.onload = () => this.onNeedsRedraw?.()
    el.src = src
    return el
  }

  /** 缺图提示(仅画淡图标,**不盖底色** — 让单元格自身填充色透出来) */
  private drawImagePlaceholder(x: number, y: number, w: number, h: number, zoom: number): void {
    const ctx = this.ctx
    const fs = Math.min(h * 0.5, 14 * zoom)
    if (fs < 6) return
    ctx.save()
    ctx.fillStyle = '#c0c6cf'
    ctx.font = `${fs}px sans-serif`
    ctx.textAlign = 'center'
    ctx.textBaseline = 'middle'
    ctx.fillText('🖼', x + w / 2, y + h / 2)
    ctx.restore()
  }

  private drawCellContent(
    cell: CellModel,
    style: CellStyle,
    row: number,
    col: number,
    x: number,
    y: number,
    w: number,
    h: number,
    zoom: number,
    effect: CellEffect | null,
  ): void {
    const ctx = this.ctx
    const pad = CELL_PADDING * zoom
    const isNumber = cell.type === 'number' || (cell.type === 'formula' && typeof cell.raw === 'number')

    // 富文本特殊处理
    if (cell.type === 'richtext' && cell.rich) {
      this.drawRichText(cell, style, x, y, w, h, zoom)
      return
    }

    const formatted = formatValue(cell.raw, style.numFmt, this.workbook.date1904)
    let text = formatted.text
    if (cell.type === 'hyperlink' && !text) text = String(cell.hyperlink ?? '')
    if (!text) return

    const color = effect?.fontColor || formatted.color ||
      (cell.type === 'hyperlink' ? this.workbook.themeColors[10] || '#0563C1' : style.font.color)
    let fontCss = fontToCss(style.font, zoom)
    if (effect?.bold && !style.font.bold) {
      fontCss = fontToCss({ ...style.font, bold: true }, zoom)
    }
    ctx.font = fontCss
    ctx.fillStyle = color

    const hAlign = resolveHAlign(style.hAlign, isNumber)
    const rotation = style.textRotation
    const availW = w - pad * 2

    // 旋转文本
    if (rotation && rotation !== 0 && rotation !== 255) {
      this.drawRotatedText(text, style, color, fontCss, x, y, w, h, zoom)
      return
    }
    if (rotation === 255) {
      this.drawVerticalText(text, color, fontCss, x, y, w, h, style)
      return
    }

    // 文本溢出: 非换行/非缩放文本若超出列宽，且相邻是空格，溢出铺过去(像 Excel)
    let clipX = x
    let clipW = w
    if (!style.wrapText && !style.shrinkToFit) {
      const tw = measureWidth(ctx, text, fontCss)
      if (tw > availW) {
        const spill = this.spillClip(row, col, x, w, tw - availW, hAlign)
        clipX = spill.x
        clipW = spill.w
      }
    }

    ctx.save()
    ctx.beginPath()
    ctx.rect(clipX, y, clipW, h)
    ctx.clip()

    let lines: string[]
    let drawFont = fontCss
    if (style.wrapText) {
      lines = wrapLines(ctx, text, fontCss, availW)
    } else if (style.shrinkToFit) {
      const scale = shrinkScale(ctx, text, fontCss, availW)
      if (scale < 1) drawFont = fontToCss({ ...style.font, size: style.font.size * scale }, zoom)
      lines = [text]
    } else {
      lines = [text]
    }
    ctx.font = drawFont

    const lineH = style.font.size * zoom * (96 / 72) * LINE_HEIGHT_FACTOR
    const totalH = lineH * lines.length
    const availH = h - 2 * pad
    // Phase 1 长文本撑高后 (2026-06-08): 文本总高度超过单元格 → 强制顶对齐, 显示文头 (跟 WPS 一致).
    // 否则默认 'bottom' 会从底部减去 totalH 算 startY, 导致 startY < y, 第一行画到格外, 用户看到的是文末.
    const overflowsCell = totalH > availH
    let startY: number
    if (style.vAlign === 'top' || overflowsCell) startY = y + pad + lineH * 0.78
    else if (style.vAlign === 'middle') startY = y + (h - totalH) / 2 + lineH * 0.78
    else startY = y + h - pad - totalH + lineH * 0.78

    ctx.textAlign = 'left'
    ctx.textBaseline = 'alphabetic'
    const indentPx = style.indent * 8 * zoom

    for (let i = 0; i < lines.length; i++) {
      const ln = lines[i]
      const lw = measureWidth(ctx, ln, drawFont)
      let tx = x + pad + indentPx
      if (hAlign === 'center') tx = x + (w - lw) / 2
      else if (hAlign === 'right') tx = x + w - pad - lw
      const ty = startY + i * lineH
      ctx.fillText(ln, tx, ty)
      this.drawTextDecoration(style, color, tx, ty, lw)
    }
    ctx.restore()
  }

  /**
   * 文本溢出可用的裁剪范围: 从本格向对齐方向延伸，吞掉相邻"空白"格(无值且无填充)。
   * 返回扩展后的 clip x/w(对齐锚点仍用原格，只是裁剪放宽)。
   */
  private spillClip(
    row: number,
    col: number,
    x: number,
    w: number,
    overflowPx: number,
    hAlign: string,
  ): { x: number; w: number } {
    if (this.merges.rangeOf(row, col)) return { x, w }
    const isBlank = (r: number, c: number): boolean => {
      if (c < 0 || c >= this.metrics.cols) return false
      if (this.merges.rangeOf(r, c)) return false
      const cell = this.sheet.cells.get(cellKey(r, c))
      if (!cell) return true
      if (cell.type !== 'empty' && cell.raw !== null && cell.raw !== '') return false // 有值 → 挡住
      const st = this.sheet.styles[cell.styleId]
      return !st || st.fill.type === 'none' // 有填充的空格也挡住(避免被覆盖)
    }
    let left = 0
    let right = 0
    if (hAlign === 'right') {
      let c = col - 1
      while (left < overflowPx && isBlank(row, c)) {
        left += this.metrics.colWidth(c)
        c--
      }
    } else if (hAlign === 'center') {
      let cl = col - 1
      let cr = col + 1
      while (left + right < overflowPx) {
        if (right <= left && isBlank(row, cr)) right += this.metrics.colWidth(cr++)
        else if (right > left && isBlank(row, cl)) left += this.metrics.colWidth(cl--)
        else if (isBlank(row, cr)) right += this.metrics.colWidth(cr++)
        else if (isBlank(row, cl)) left += this.metrics.colWidth(cl--)
        else break
      }
    } else {
      // left / general(文本) / fill / justify → 向右溢出
      let c = col + 1
      while (right < overflowPx && isBlank(row, c)) {
        right += this.metrics.colWidth(c)
        c++
      }
    }
    return { x: x - left, w: w + left + right }
  }

  private drawRichText(cell: CellModel, style: CellStyle, x: number, y: number, w: number, h: number, zoom: number): void {
    const ctx = this.ctx
    const pad = CELL_PADDING * zoom
    const runs = cell.rich || []
    const baseFont = style.font
    const hAlign = resolveHAlign(style.hAlign, false)
    const indentPx = (style.indent || 0) * 8 * zoom
    const availW = w - 2 * pad - (hAlign === 'left' ? indentPx : 0)

    // shrinkToFit(且非 wrap):按单行总宽算统一缩放,塞进列宽(跟普通文本路径同档)
    let scale = 1
    if (style.shrinkToFit && !style.wrapText && availW > 0) {
      let totalW = 0
      for (let i = 0; i < runs.length; i++) { ctx.font = fontToCss({ ...baseFont, ...runs[i].font } as any, zoom); totalW += ctx.measureText(runs[i].text).width }
      if (totalW > availW && totalW > 0) scale = Math.max(0.3, availW / totalW)
    }

    // 每个 run 预算:字体串 + 颜色 + 下划线/删除线 + 缩放后字号(px,画装饰线用)
    type Seg = { text: string; fontCss: string; color: string; ul: boolean; st: boolean; szPx: number }
    const fontCache: string[] = []; const colorOf: string[] = []; const ulOf: boolean[] = []; const stOf: boolean[] = []; const szPxOf: number[] = []
    for (let i = 0; i < runs.length; i++) {
      const f = { ...baseFont, ...runs[i].font }
      const szPt = (f.size ?? baseFont.size) * scale
      fontCache[i] = fontToCss({ ...f, size: szPt } as any, zoom)
      colorOf[i] = (runs[i].font?.color as string) || baseFont.color
      ulOf[i] = !!f.underline; stOf[i] = !!f.strike; szPxOf[i] = szPt * zoom * (96 / 72)
    }
    const mkSeg = (ch: string, i: number): Seg => ({ text: ch, fontCss: fontCache[i], color: colorOf[i], ul: ulOf[i], st: stOf[i], szPx: szPxOf[i] })

    // 排版成行:wrapText 逐字符按列宽折行(保留各 run 字体/颜色/装饰);否则单行
    const lines: Seg[][] = []
    if (style.wrapText && availW > 0) {
      let line: Seg[] = []
      let lineW = 0
      const flush = () => { lines.push(line); line = []; lineW = 0 }
      for (let i = 0; i < runs.length; i++) {
        ctx.font = fontCache[i]
        for (const ch of runs[i].text) {
          if (ch === '\n') { flush(); continue }
          const cw = ctx.measureText(ch).width
          if (lineW + cw > availW && lineW > 0) flush()
          const last = line[line.length - 1]
          if (last && last.fontCss === fontCache[i] && last.color === colorOf[i] && last.ul === ulOf[i] && last.st === stOf[i]) last.text += ch
          else line.push(mkSeg(ch, i))
          lineW += cw
        }
      }
      flush()
    } else {
      lines.push(runs.map((_, i) => ({ ...mkSeg(runs[i].text, i) })))
    }

    // 垂直对齐 + 溢出顶对齐(跟普通文本 drawText 一致:超出格高 → 顶对齐显示文头,WPS 行为)
    const lineH = baseFont.size * scale * zoom * (96 / 72) * LINE_HEIGHT_FACTOR
    const totalH = lineH * lines.length
    const overflowsCell = totalH > h - 2 * pad
    let startY: number
    if (style.vAlign === 'top' || overflowsCell) startY = y + pad + lineH * 0.78
    else if (style.vAlign === 'middle') startY = y + (h - totalH) / 2 + lineH * 0.78
    else startY = y + h - pad - totalH + lineH * 0.78

    ctx.save()
    ctx.beginPath()
    ctx.rect(x, y, w, h)
    ctx.clip()
    ctx.textBaseline = 'alphabetic'
    ctx.textAlign = 'left'
    for (let li = 0; li < lines.length; li++) {
      const segs = lines[li]
      let lineW = 0
      for (const s of segs) { ctx.font = s.fontCss; lineW += ctx.measureText(s.text).width }
      let tx = x + pad + indentPx
      if (hAlign === 'center') tx = x + (w - lineW) / 2
      else if (hAlign === 'right') tx = x + w - pad - lineW
      const ty = startY + li * lineH
      for (const s of segs) {
        ctx.font = s.fontCss
        ctx.fillStyle = s.color
        ctx.fillText(s.text, tx, ty)
        const sw = ctx.measureText(s.text).width
        if (s.ul || s.st) { // 逐 run 下划线/删除线
          ctx.strokeStyle = s.color
          ctx.lineWidth = Math.max(1, s.szPx / 14)
          if (s.ul) { ctx.beginPath(); ctx.moveTo(tx, ty + 2); ctx.lineTo(tx + sw, ty + 2); ctx.stroke() }
          if (s.st) { const sy = ty - s.szPx * 0.3; ctx.beginPath(); ctx.moveTo(tx, sy); ctx.lineTo(tx + sw, sy); ctx.stroke() }
        }
        tx += sw
        if (tx > x + w + 50) break // 非折行的超长单行:画出格外即停
      }
    }
    ctx.restore()
  }

  private drawRotatedText(text: string, style: CellStyle, color: string, fontCss: string, x: number, y: number, w: number, h: number, _zoom: number): void {
    const ctx = this.ctx
    ctx.save()
    ctx.beginPath()
    ctx.rect(x, y, w, h)
    ctx.clip()
    ctx.translate(x + w / 2, y + h / 2)
    // Excel: 正角度逆时针。canvas rotate 顺时针为正，故取负。
    ctx.rotate((-style.textRotation * Math.PI) / 180)
    ctx.font = fontCss
    ctx.fillStyle = color
    ctx.textAlign = 'center'
    ctx.textBaseline = 'middle'
    ctx.fillText(text, 0, 0)
    ctx.restore()
  }

  private drawVerticalText(text: string, color: string, fontCss: string, x: number, y: number, w: number, h: number, style: CellStyle): void {
    const ctx = this.ctx
    ctx.save()
    ctx.beginPath()
    ctx.rect(x, y, w, h)
    ctx.clip()
    ctx.font = fontCss
    ctx.fillStyle = color
    ctx.textAlign = 'center'
    ctx.textBaseline = 'top'
    const chars = [...text]
    const lineH = parseFloat(fontCss) * 1.05
    // 竖排也尊重 vAlign:字符串总高 vs 格高 → top/middle/bottom 起点(超高则顶对齐)
    const totalH = chars.length * lineH
    let ty = y + 3
    if (totalH <= h - 6) {
      if (style.vAlign === 'middle') ty = y + (h - totalH) / 2
      else if (style.vAlign === 'bottom') ty = y + h - totalH - 3
    }
    for (const ch of chars) {
      ctx.fillText(ch, x + w / 2, ty)
      ty += lineH
      if (ty > y + h) break
    }
    ctx.restore()
  }

  private drawTextDecoration(style: CellStyle, color: string, tx: number, ty: number, width: number): void {
    const ctx = this.ctx
    if (!style.font.underline && !style.font.strike) return
    ctx.save()
    ctx.strokeStyle = color
    ctx.lineWidth = 1
    if (style.font.underline) {
      ctx.beginPath()
      ctx.moveTo(tx, Math.round(ty + 2) + 0.5)
      ctx.lineTo(tx + width, Math.round(ty + 2) + 0.5)
      ctx.stroke()
    }
    if (style.font.strike) {
      const sy = Math.round(ty - style.font.size * 0.35) + 0.5
      ctx.beginPath()
      ctx.moveTo(tx, sy)
      ctx.lineTo(tx + width, sy)
      ctx.stroke()
    }
    ctx.restore()
  }

  // ---------------- 表头 ----------------
  private drawHeaders(panes: Pane[], view: ViewState): void {
    const ctx = this.ctx
    const hw = this.metrics.rowHeaderWidth
    const hh = this.metrics.colHeaderHeight
    void panes
    const frozenCols = this.freeze.frozenCols
    const frozenRows = this.freeze.frozenRows
    const fw = this.freeze.frozenWidth
    const fh = this.freeze.frozenHeight

    // 滚动区可视行列(可超出数据范围以铺满表头，模拟 Excel)
    const scColStart = Math.max(frozenCols, this.metrics.colAt(fw + view.scrollX))
    const scColEnd = this.metrics.colAt(view.width - hw + view.scrollX) + 1
    const scRowStart = Math.max(frozenRows, this.metrics.rowAt(fh + view.scrollY))
    const scRowEnd = this.metrics.rowAt(view.height - hh + view.scrollY) + 1

    // 列表头
    ctx.save()
    ctx.beginPath()
    ctx.rect(hw, 0, view.width - hw, hh)
    ctx.clip()
    // 滚动列
    this.drawColHeaderRange(scColStart, scColEnd, view.scrollX, hw, hh)
    ctx.restore()
    // 冻结列(覆盖在滚动列之上)
    if (frozenCols > 0) {
      ctx.save()
      ctx.beginPath()
      ctx.rect(hw, 0, this.freeze.frozenWidth, hh)
      ctx.clip()
      this.drawColHeaderRange(0, frozenCols - 1, 0, hw, hh)
      ctx.restore()
    }

    // 行表头
    ctx.save()
    ctx.beginPath()
    ctx.rect(0, hh, hw, view.height - hh)
    ctx.clip()
    this.drawRowHeaderRange(scRowStart, scRowEnd, view.scrollY, hw, hh)
    ctx.restore()
    if (frozenRows > 0) {
      ctx.save()
      ctx.beginPath()
      ctx.rect(0, hh, hw, this.freeze.frozenHeight)
      ctx.clip()
      this.drawRowHeaderRange(0, frozenRows - 1, 0, hw, hh)
      ctx.restore()
    }

    // 左上角
    ctx.fillStyle = this.theme.headerBg
    ctx.fillRect(0, 0, hw, hh)
    ctx.strokeStyle = this.theme.headerLine
    ctx.lineWidth = 1
    ctx.strokeRect(0.5, 0.5, hw, hh)
  }

  private drawColHeaderRange(start: number, end: number, scrollX: number, hw: number, hh: number): void {
    const ctx = this.ctx
    ctx.textAlign = 'center'
    ctx.textBaseline = 'middle'
    ctx.font = '11px Calibri, sans-serif'
    for (let c = start; c <= end; c++) {
      const x = hw + this.metrics.colLeft(c) - scrollX
      const w = this.metrics.colWidth(c)
      if (w <= 0) continue
      ctx.fillStyle = this.theme.headerBg
      ctx.fillRect(x, 0, w, hh)
      ctx.strokeStyle = this.theme.headerLine
      ctx.lineWidth = 1
      ctx.beginPath()
      ctx.moveTo(Math.round(x) + 0.5, 0)
      ctx.lineTo(Math.round(x) + 0.5, hh)
      ctx.moveTo(x, hh - 0.5)
      ctx.lineTo(x + w, hh - 0.5)
      ctx.stroke()
      ctx.fillStyle = this.theme.headerText
      ctx.fillText(colIndexToLetters(c), x + w / 2, hh / 2 + 1)
    }
  }

  private drawRowHeaderRange(start: number, end: number, scrollY: number, hw: number, hh: number): void {
    const ctx = this.ctx
    ctx.textAlign = 'center'
    ctx.textBaseline = 'middle'
    ctx.font = '11px Calibri, sans-serif'
    for (let r = start; r <= end; r++) {
      const y = hh + this.metrics.rowTop(r) - scrollY
      const h = this.metrics.rowHeight(r)
      if (h <= 0) continue
      ctx.fillStyle = this.theme.headerBg
      ctx.fillRect(0, y, hw, h)
      ctx.strokeStyle = this.theme.headerLine
      ctx.lineWidth = 1
      ctx.beginPath()
      ctx.moveTo(0, Math.round(y) + 0.5)
      ctx.lineTo(hw, Math.round(y) + 0.5)
      ctx.moveTo(hw - 0.5, y)
      ctx.lineTo(hw - 0.5, y + h)
      ctx.stroke()
      ctx.fillStyle = this.theme.headerText
      ctx.fillText(String(r + 1), hw / 2, y + h / 2 + 1)
    }
  }

  private drawFreezeLines(layout: { rowHeaderWidth: number; colHeaderHeight: number }): void {
    const ctx = this.ctx
    if (this.freeze.frozenCols === 0 && this.freeze.frozenRows === 0) return
    ctx.save()
    ctx.strokeStyle = '#9AA4AE'
    ctx.lineWidth = 1
    if (this.freeze.frozenCols > 0) {
      const x = layout.rowHeaderWidth + this.freeze.frozenWidth + 0.5
      ctx.beginPath()
      ctx.moveTo(x, 0)
      ctx.lineTo(x, ctx.canvas.height)
      ctx.stroke()
    }
    if (this.freeze.frozenRows > 0) {
      const y = layout.colHeaderHeight + this.freeze.frozenHeight + 0.5
      ctx.beginPath()
      ctx.moveTo(0, y)
      ctx.lineTo(ctx.canvas.width, y)
      ctx.stroke()
    }
    ctx.restore()
  }

  private styleOf(cell: CellModel): CellStyle {
    const base = this.sheet.styles[cell.styleId] ?? defaultCellStyle()
    const editable = this.isEditableFn(cell.row, cell.col)
    let out: CellStyle = base
    // ① cellStyle 钩子 (传入 ctx.editable;旧 (cell, pos) => ... 签名兼容,第 3 入参可选)
    if (this.cellStyleHook) {
      const over = this.cellStyleHook(cell, { row: cell.row, col: cell.col }, { editable })
      if (over) out = mergeStyleOverride(out, over)
    }
    // ② 只读视觉钩子 (Phase C 2026-06-08): 仅在该格 !editable 且配了 readOnlyCellStyle 时套
    if (!editable && this.readOnlyStyleHook) {
      const ro = this.readOnlyStyleHook
      let roOver: CellStyleOverride | void = undefined
      if (ro === true) {
        // 内置默认: 浅灰底, 跟工具栏背景一致, 不抢眼
        roOver = { fill: { type: 'solid', fgColor: '#f5f7fa' } }
      } else if (typeof ro === 'function') {
        roOver = ro(cell, { row: cell.row, col: cell.col }, { editable })
      } else if (ro && typeof ro === 'object') {
        roOver = ro
      }
      if (roOver) out = mergeStyleOverride(out, roOver)
    }
    return out
  }
}

function defaultCellStyle(): CellStyle {
  return {
    font: { name: 'Calibri', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#000000' },
    fill: { type: 'none' },
    borders: {},
    hAlign: 'general',
    vAlign: 'bottom',
    wrapText: false,
    shrinkToFit: false,
    textRotation: 0,
    indent: 0,
    numFmt: 'General',
  }
}

// ---------------- 小工具 ----------------
function withAlpha(color: string, alpha: number): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(color)
  if (!m) return color
  const n = parseInt(m[1], 16)
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${alpha})`
}

/**
 * [start,end] 减去 gaps(合并区内部区间,可重叠/乱序)的补集 → 要画的实线段。
 * 纯函数,便于单测。合并区内部的网格线靠此被"挖空"。
 */
export function gridSegments(start: number, end: number, gaps: Array<[number, number]>): Array<[number, number]> {
  if (!gaps.length) return start < end ? [[start, end]] : []
  const sorted = [...gaps].sort((a, b) => a[0] - b[0])
  const out: Array<[number, number]> = []
  let cur = start
  for (const [gs, ge] of sorted) {
    const s = Math.max(gs, start)
    const e = Math.min(ge, end)
    if (s > cur) out.push([cur, s]) // 空隙前的实线段
    if (e > cur) cur = e // 推进越过本合并区
  }
  if (cur < end) out.push([cur, end])
  return out
}

/**
 * 往当前 path 里加 [start,end] 网格线段,跳过 gaps(合并区内部)。
 * vertical=true: 竖线,fixed=x,start/end 为 y;否则横线,fixed=y,start/end 为 x。
 */
function strokeGapped(
  ctx: CanvasRenderingContext2D,
  fixed: number,
  start: number,
  end: number,
  gaps: Array<[number, number]>,
  vertical: boolean,
): void {
  for (const [a, b] of gridSegments(start, end, gaps)) {
    if (b - a <= 0.0001) continue
    if (vertical) {
      ctx.moveTo(fixed, a)
      ctx.lineTo(fixed, b)
    } else {
      ctx.moveTo(a, fixed)
      ctx.lineTo(b, fixed)
    }
  }
}

/** 在单元格内画迷你图(折线/柱/盈亏)。 */
function drawSparkline(ctx: CanvasRenderingContext2D, sp: Sparkline, x: number, y: number, w: number, h: number): void {
  const pad = 2
  const ix = x + pad
  const iy = y + pad
  const iw = w - pad * 2
  const ih = h - pad * 2
  if (iw < 4 || ih < 3) return
  const vals = sp.values
  const nums = vals.filter((v): v is number => v !== null)
  if (!nums.length) return
  const min = Math.min(...nums)
  const max = Math.max(...nums)
  const color = sp.color || '#376092'
  const negColor = sp.negativeColor || '#D00000'
  const n = vals.length

  ctx.save()
  if (sp.type === 'winloss') {
    const bw = iw / n
    const mid = iy + ih / 2
    const bh = ih * 0.4
    for (let i = 0; i < n; i++) {
      const v = vals[i]
      if (v == null || v === 0) continue
      ctx.fillStyle = v > 0 ? color : negColor
      ctx.fillRect(ix + i * bw + bw * 0.15, v > 0 ? mid - bh : mid, bw * 0.7, bh)
    }
  } else if (sp.type === 'column') {
    const hi = Math.max(max, 0)
    const lo = Math.min(min, 0)
    const span = hi - lo || 1
    const zeroY = iy + ih * (hi / span)
    const bw = iw / n
    for (let i = 0; i < n; i++) {
      const v = vals[i]
      if (v == null) continue
      ctx.fillStyle = v < 0 ? negColor : color
      const vy = iy + ih * ((hi - v) / span)
      ctx.fillRect(ix + i * bw + bw * 0.15, Math.min(vy, zeroY), bw * 0.7, Math.max(1, Math.abs(vy - zeroY)))
    }
  } else {
    // line
    const range = max - min || 1
    ctx.strokeStyle = color
    ctx.lineWidth = 1
    ctx.beginPath()
    let started = false
    for (let i = 0; i < n; i++) {
      const v = vals[i]
      if (v == null) {
        started = false
        continue
      }
      const px = ix + (n === 1 ? iw / 2 : (i / (n - 1)) * iw)
      const py = iy + ih - ((v - min) / range) * ih
      if (!started) {
        ctx.moveTo(px, py)
        started = true
      } else {
        ctx.lineTo(px, py)
      }
    }
    ctx.stroke()
  }
  ctx.restore()
}

function drawIcon(ctx: CanvasRenderingContext2D, icon: { setName: string; level: number; count: number }, x: number, y: number, h: number): void {
  const cy = y + h / 2
  const cx = x + 7
  const r = 5
  ctx.save()
  if (icon.setName.includes('TrafficLights') || icon.setName.includes('Signs') || icon.count === 3 && icon.setName.includes('Symbols')) {
    const colors = ['#D63B3B', '#E8B53A', '#5B9F4E']
    ctx.fillStyle = colors[Math.min(icon.level, 2)]
    ctx.beginPath()
    ctx.arc(cx, cy, r, 0, Math.PI * 2)
    ctx.fill()
  } else if (icon.setName.includes('Arrows')) {
    // 箭头: 低→下(红) 中→平(黄) 高→上(绿)
    const colors = ['#D63B3B', '#E8B53A', '#5B9F4E', '#5B9F4E', '#5B9F4E']
    ctx.strokeStyle = colors[Math.min(icon.level, colors.length - 1)]
    ctx.fillStyle = ctx.strokeStyle
    ctx.lineWidth = 2
    const dir = icon.level === 0 ? Math.PI / 2 : icon.level >= icon.count - 1 ? -Math.PI / 2 : 0
    ctx.translate(cx, cy)
    ctx.rotate(dir)
    ctx.beginPath()
    ctx.moveTo(-4, 0)
    ctx.lineTo(4, 0)
    ctx.moveTo(1, -3)
    ctx.lineTo(4, 0)
    ctx.lineTo(1, 3)
    ctx.stroke()
  } else {
    // 通用: 用色阶圆点
    const t = icon.count > 1 ? icon.level / (icon.count - 1) : 1
    ctx.fillStyle = `rgb(${Math.round(214 - 120 * t)},${Math.round(59 + 100 * t)},${59})`
    ctx.beginPath()
    ctx.arc(cx, cy, r, 0, Math.PI * 2)
    ctx.fill()
  }
  ctx.restore()
}
