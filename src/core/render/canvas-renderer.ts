/**
 * Canvas 主渲染器。组合几何/合并/冻结/条件格式/格式引擎，按可视区重绘。
 * 渲染顺序(每个 pane 内): 网格线 → 填充/条件背景 → 数据条 → 边框 → 文本/图标 → 筛选按钮。
 * 表头(行号/列字母)最后绘制，覆盖在最上层。
 */
import type { CellModel, CellStyle, CellStyleFn, MergeRange, SheetModel, Sparkline, WorkbookModel } from '../model/types'
import { cellKey } from '../model/types'
import { type ViewerTheme, mergeTheme } from './theme'
import { GridMetrics, colIndexToLetters } from '../layout/grid-metrics'
import { MergeIndex } from '../layout/merges'
import { computeFreeze, type FreezeGeometry } from '../layout/freeze'
import { computeViewport, type Pane } from '../layout/viewport'
import { ConditionalEngine, type CellEffect } from './conditional'
import { formatValue } from '../format/number-format'
import { paintFill } from './fills'
import { drawEdge } from './borders'
import {
  fontToCss, measureWidth, resolveHAlign, shrinkScale, wrapLines,
  LINE_HEIGHT_FACTOR, CELL_PADDING,
} from './text'
import { autoFitRowHeights } from '../layout/autofit'
import { drawFilterButton, isFilterHeader } from './autofilter'

export interface RendererOptions {
  theme?: Partial<ViewerTheme>
  cellStyle?: CellStyleFn
}

export interface ViewState {
  scrollX: number
  scrollY: number
  width: number // css px
  height: number
  zoom: number
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

  constructor(
    private canvas: HTMLCanvasElement,
    private sheet: SheetModel,
    private workbook: WorkbookModel,
    zoom = 1,
    opts?: RendererOptions,
  ) {
    this.theme = mergeTheme(opts?.theme)
    this.cellStyleHook = opts?.cellStyle
    const ctx = canvas.getContext('2d')
    if (!ctx) throw new Error('无法获取 canvas 2d context')
    this.ctx = ctx
    // 先做自动行高(撑高换行溢出的行)，再据此建几何
    autoFitRowHeights(sheet, workbook, ctx)
    this.metrics = new GridMetrics(sheet, zoom)
    this.merges = new MergeIndex(sheet)
    this.freeze = computeFreeze(sheet, this.metrics)
    this.cond = new ConditionalEngine(sheet)
    for (const sp of sheet.sparklines) this.sparklineIndex.set(cellKey(sp.row, sp.col), sp)
  }

  /** 改变缩放: 重建几何(列宽行高表头按 zoom 缩放)。合并/条件格式无需重建。 */
  setZoom(zoom: number): void {
    if (zoom === this.metrics.zoom) return
    this.metrics = new GridMetrics(this.sheet, zoom)
    this.freeze = computeFreeze(this.sheet, this.metrics)
  }

  /** 内容总尺寸(含表头)，给外层滚动容器用 */
  get contentWidth(): number {
    return this.metrics.rowHeaderWidth + this.metrics.totalWidth
  }
  get contentHeight(): number {
    return this.metrics.colHeaderHeight + this.metrics.totalHeight
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

  /** 屏幕坐标 → 单元格(0-based)。落在表头或越界返回 null。 */
  cellAtScreen(view: ViewState, x: number, y: number): { row: number; col: number } | null {
    const hw = this.metrics.rowHeaderWidth
    const hh = this.metrics.colHeaderHeight
    if (x < hw || y < hh) return null
    const fw = this.freeze.frozenWidth
    const fh = this.freeze.frozenHeight
    const cx = x < hw + fw ? x - hw : x - hw + view.scrollX
    const cy = y < hh + fh ? y - hh : y - hh + view.scrollY
    const col = Math.min(Math.max(this.metrics.colAt(cx), 0), this.metrics.cols - 1)
    const row = Math.min(Math.max(this.metrics.rowAt(cy), 0), this.metrics.rows - 1)
    return { row, col }
  }

  /** 屏幕 y → 行(用于点行号选整行);落在列表头返回 -1 */
  rowAtScreen(view: ViewState, y: number): number {
    const hh = this.metrics.colHeaderHeight
    if (y < hh) return -1
    const fh = this.freeze.frozenHeight
    const cy = y < hh + fh ? y - hh : y - hh + view.scrollY
    return Math.min(Math.max(this.metrics.rowAt(cy), 0), this.metrics.rows - 1)
  }
  /** 屏幕 x → 列(用于点列标选整列);落在行表头返回 -1 */
  colAtScreen(view: ViewState, x: number): number {
    const hw = this.metrics.rowHeaderWidth
    if (x < hw) return -1
    const fw = this.freeze.frozenWidth
    const cx = x < hw + fw ? x - hw : x - hw + view.scrollX
    return Math.min(Math.max(this.metrics.colAt(cx), 0), this.metrics.cols - 1)
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
    this.metrics = new GridMetrics(this.sheet, this.metrics.zoom)
    this.freeze = computeFreeze(this.sheet, this.metrics)
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

  /** 单元格的显示文本(套数字格式后);空返回 '' */
  cellText(row: number, col: number): string {
    const cell = this.sheet.cells.get(cellKey(row, col))
    if (!cell || cell.type === 'empty') return ''
    if (cell.type === 'richtext' && cell.rich) return cell.rich.map((r) => r.text).join('')
    return formatValue(cell.raw, this.styleOf(cell).numFmt, this.workbook.date1904).text
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
    const { width, height, zoom } = view
    this.dpr = window.devicePixelRatio || 1
    // 调整 canvas 像素尺寸(高清)
    const pw = Math.round(width * this.dpr)
    const ph = Math.round(height * this.dpr)
    if (this.canvas.width !== pw || this.canvas.height !== ph) {
      this.canvas.width = pw
      this.canvas.height = ph
    }
    const ctx = this.ctx
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0)
    ctx.clearRect(0, 0, width, height)
    ctx.fillStyle = '#FFFFFF'
    ctx.fillRect(0, 0, width, height)

    const layout = computeViewport(this.metrics, this.freeze, view.scrollX, view.scrollY, width, height)

    for (const pane of layout.panes) {
      this.drawPane(pane, zoom)
    }

    this.drawHeaders(layout.panes, view)
    this.drawFreezeLines(layout)
    this.drawPageBreaks(view)
    this.drawSelection(view)
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
    // 单格选区不填充(像 Excel 的活动单元格),多格才铺淡蓝
    const single = sel.top === sel.bottom && sel.left === sel.right
    if (!single) {
      ctx.fillStyle = this.theme.selFill
      ctx.fillRect(x, y, w, h)
    }
    ctx.strokeStyle = this.theme.selBorder
    ctx.lineWidth = 2
    ctx.strokeRect(Math.round(x) + 1, Math.round(y) + 1, Math.round(w) - 2, Math.round(h) - 2)
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

    // 1. 网格线
    if (this.sheet.showGridLines) {
      ctx.strokeStyle = this.theme.gridLine
      ctx.lineWidth = 1
      ctx.beginPath()
      for (let c = Math.max(0, gridC0); c <= gridC1; c++) {
        const x = Math.round(sx(c)) + 0.5
        ctx.moveTo(x, pane.clipY)
        ctx.lineTo(x, pane.clipY + pane.clipH)
      }
      for (let r = Math.max(0, gridR0); r <= gridR1; r++) {
        const y = Math.round(sy(r)) + 0.5
        ctx.moveTo(pane.clipX, y)
        ctx.lineTo(pane.clipX + pane.clipW, y)
      }
      ctx.stroke()
    }

    // 2. 合并区(锚点可能在可视区外，需单独扫描)
    const coveredAnchorsDrawn = new Set<string>()
    for (const m of this.sheet.merges) {
      if (m.bottom < pane.rowStart || m.top > pane.rowEnd || m.right < pane.colStart || m.left > pane.colEnd) continue
      const x = sx(m.left)
      const y = sy(m.top)
      const w = this.metrics.colLeft(m.right + 1) - this.metrics.colLeft(m.left)
      const h = this.metrics.rowTop(m.bottom + 1) - this.metrics.rowTop(m.top)
      const cell = this.sheet.cells.get(cellKey(m.top, m.left))
      this.paintCellBox(cell, m.top, m.left, x, y, w, h, zoom)
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
      drawEdge(ctx, b.top, x, y, x + w, y)
      drawEdge(ctx, b.bottom, x, y + h, x + w, y + h)
      drawEdge(ctx, b.left, x, y, x, y + h)
      drawEdge(ctx, b.right, x + w, y, x + w, y + h)
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

    // 内容
    if (cell && cell.type !== 'empty' && style) {
      this.drawCellContent(cell, style, row, col, contentX, y, contentW, h, zoom, effect)
    }

    // 自动筛选下拉
    if (isFilterHeader(this.sheet.autoFilterRange, row, col)) {
      drawFilterButton(ctx, x, y, w, h)
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
    let startY: number
    if (style.vAlign === 'top') startY = y + pad + lineH * 0.78
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

    // 预测总宽(用于水平对齐)
    let totalW = 0
    const fontCache: string[] = []
    for (let i = 0; i < runs.length; i++) {
      const f = fontToCss({ ...baseFont, ...runs[i].font } as any, zoom)
      fontCache[i] = f
      ctx.font = f
      totalW += ctx.measureText(runs[i].text).width
    }

    // 水平起点(尊重 hAlign;富文本默认按文本左对齐)
    const hAlign = resolveHAlign(style.hAlign, false)
    let tx = x + pad
    if (hAlign === 'center') tx = x + (w - totalW) / 2
    else if (hAlign === 'right') tx = x + w - pad - totalW

    // 垂直基线(尊重 vAlign)
    const asc = baseFont.size * zoom * (96 / 72) * 0.72
    const lineH = baseFont.size * zoom * (96 / 72) * LINE_HEIGHT_FACTOR
    let ty: number
    if (style.vAlign === 'top') ty = y + pad + asc
    else if (style.vAlign === 'middle') ty = y + (h - lineH) / 2 + asc
    else ty = y + h - pad - lineH + asc

    ctx.save()
    ctx.beginPath()
    ctx.rect(x, y, w, h)
    ctx.clip()
    ctx.textBaseline = 'alphabetic'
    ctx.textAlign = 'left'
    for (let i = 0; i < runs.length; i++) {
      ctx.font = fontCache[i]
      ctx.fillStyle = (runs[i].font?.color as string) || baseFont.color
      ctx.fillText(runs[i].text, tx, ty)
      tx += ctx.measureText(runs[i].text).width
      if (tx > x + w + 50) break
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

  private drawVerticalText(text: string, color: string, fontCss: string, x: number, y: number, w: number, h: number, _style: CellStyle): void {
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
    let ty = y + 3
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
    const base = this.sheet.styles[cell.styleId]
    if (!this.cellStyleHook) return base
    const over = this.cellStyleHook(cell, { row: cell.row, col: cell.col })
    return over ? applyStyleOverride(base, over) : base
  }
}

/** 合并 cellStyle 钩子返回的部分样式: font/fill/borders 浅合并,其余覆盖 */
function applyStyleOverride(base: CellStyle, over: Partial<CellStyle>): CellStyle {
  return {
    ...base,
    ...over,
    font: over.font ? { ...base.font, ...over.font } : base.font,
    fill: over.fill ? { ...base.fill, ...over.fill } : base.fill,
    borders: over.borders ? { ...base.borders, ...over.borders } : base.borders,
  }
}

// ---------------- 小工具 ----------------
function withAlpha(color: string, alpha: number): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(color)
  if (!m) return color
  const n = parseInt(m[1], 16)
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${alpha})`
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
