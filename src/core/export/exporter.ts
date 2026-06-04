/**
 * WorkbookExporter(框架无关)—— 把"工作簿 → 图片 / PDF / 打印"的编排逻辑从框架壳里抽出。
 *
 * 只依赖一个 host 取数器(取工作簿 / 活动表索引 / 复用当前表 renderer / 渲染配置 / 文件名),
 * 不碰任何框架反应式;Vue / React 壳各自实现 host 即可复用整套导出。
 *
 * 范围优先级: opts.range > pageSetup.printArea > 整表。当前表复用 live renderer,
 * 其它表临时建离屏 renderer(同一套 RendererOptions 保证主题/cellStyle 一致)。
 */
import type * as EChartsNS from 'echarts'
import type { ChartSpec, ImageAnchor, MergeRange, SheetModel, WorkbookModel } from '../model/types'
import { CanvasRenderer, type RendererOptions } from '../render/canvas-renderer'
import { GridMetrics } from '../layout/grid-metrics'
import { anchorRect } from '../overlay/anchor'
import { chartToOption } from '../overlay/chart-mapper'
import { loadECharts } from '../overlay/echarts-loader'
import { canvasToBlob, downloadBlob, loadImage } from './raster'
import { compositeOverlays, type ExportDecorations } from './composite'
import { exportToPdf, type ExportSheetImage } from './pdf'
import { exportToVectorPdf, type VectorSheet } from './vector-pdf'
import { printSheets } from './print'
import type { ExportTarget, ImageExportOptions, PdfExportOptions, PrintOptions } from './types'

/** 导出取数器: 壳提供当前工作簿 / 活动表 / 复用渲染器 / 渲染配置 / 文件名 */
export interface ExporterHost {
  getWorkbook(): WorkbookModel | null
  /** 活动表索引(决定 'active' 目标 + 哪张表能复用 live renderer) */
  getActiveIndex(): number
  /** 当前活动表的在用 renderer(复用以省一次离屏构建);非活动表传 null 即可 */
  getLiveRenderer(): CanvasRenderer | null
  /** 离屏渲染其它表时用的渲染配置(主题/cellStyle),须与 live 一致 */
  getRendererOpts(): RendererOptions
  /** 文件名(下载默认名;缺省回退表名 / 'workbook') */
  getFileName(): string | undefined
}

export class WorkbookExporter {
  constructor(private host: ExporterHost) {}

  /** target → 工作表索引列表 */
  private resolveTargets(target: ExportTarget = 'active'): number[] {
    const wb = this.host.getWorkbook()
    if (!wb) return []
    if (target === 'all') return wb.sheets.map((_, i) => i).filter((i) => wb.sheets[i].state === 'visible')
    if (target === 'active') return [this.host.getActiveIndex()]
    if (typeof target === 'number') return [target]
    return target.filter((i) => wb.sheets[i])
  }

  /** 取某表的 renderer: 活动表复用 live,其它表离屏新建 */
  private rendererFor(sheetIdx: number, s: SheetModel, wb: WorkbookModel): CanvasRenderer {
    const live = this.host.getLiveRenderer()
    if (sheetIdx === this.host.getActiveIndex() && live) return live
    return new CanvasRenderer(document.createElement('canvas'), s, wb, 1, this.host.getRendererOpts())
  }

  private baseName(): string {
    const wb = this.host.getWorkbook()
    const fallback = wb?.sheets[this.host.getActiveIndex()]?.name || 'workbook'
    return (this.host.getFileName() || fallback).replace(/\.[^.]+$/, '')
  }

  /** 离屏渲染一个图表为 dataURL(供非当前表 / 统一合成);echarts 不可用返回 null */
  private async chartDataUrl(spec: ChartSpec, metrics: GridMetrics): Promise<string | null> {
    let echarts: typeof EChartsNS
    try {
      echarts = await loadECharts()
    } catch {
      return null
    }
    const rect = anchorRect(metrics, spec.anchor)
    const div = document.createElement('div')
    div.style.cssText = `position:fixed;left:-10000px;top:0;width:${Math.max(80, Math.round(rect.width))}px;height:${Math.max(60, Math.round(rect.height))}px`
    document.body.appendChild(div)
    const inst = echarts.init(div)
    try {
      inst.setOption(chartToOption(spec))
      return inst.getDataURL({ pixelRatio: 2, backgroundColor: '#fff' })
    } catch {
      return null
    } finally {
      inst.dispose()
      div.remove()
    }
  }

  /** 收集一个工作表的叠加层装饰(图片/图表/形状),供合成到导出底图 */
  private async collectDecorations(s: SheetModel, metrics: GridMetrics): Promise<ExportDecorations> {
    const images: { source: CanvasImageSource; anchor: ImageAnchor }[] = []
    for (const anchor of s.images) {
      if (!anchor.src) continue
      try {
        images.push({ source: await loadImage(anchor.src), anchor })
      } catch {
        /* 单张图加载失败跳过 */
      }
    }
    const charts: { source: CanvasImageSource; anchor: ImageAnchor }[] = []
    for (const chart of s.charts) {
      const url = await this.chartDataUrl(chart, metrics)
      if (!url) continue
      try {
        charts.push({ source: await loadImage(url), anchor: chart.anchor })
      } catch {
        /* 跳过 */
      }
    }
    return { images, charts, shapes: s.shapes }
  }

  /**
   * 为一个工作表生成合成底图(格子 + 图片/图表/形状)。
   * withTitles(PDF/打印): 应用 pageSetup 打印标题行(抽出标题条 + 正文剔除标题行)与缩放。
   */
  private async buildSheetImage(sheetIdx: number, opts: PdfExportOptions, withTitles = false): Promise<ExportSheetImage | null> {
    const wb = this.host.getWorkbook()
    const s = wb?.sheets[sheetIdx]
    if (!wb || !s) return null
    const r = this.rendererFor(sheetIdx, s, wb)

    const ps = s.pageSetup
    const dim = s.dimension
    const full: MergeRange = { top: 0, left: 0, bottom: Math.max(0, dim.rows - 1), right: Math.max(0, dim.cols - 1) }
    const range = opts.range ?? ps?.printArea ?? full

    // 打印标题行/列: 标题在正文上方/左侧时,抽出标题条,正文起点相应内移(避免重复)
    let titleRows: [number, number] | null = null
    let titleCols: [number, number] | null = null
    let bodyTop = range.top
    let bodyLeft = range.left
    if (withTitles && ps?.printTitleRows) {
      const [a, b] = ps.printTitleRows
      if (b >= a && a <= range.top && b < range.bottom) {
        titleRows = [a, b]
        bodyTop = Math.max(range.top, b + 1)
      }
    }
    if (withTitles && ps?.printTitleCols) {
      const [a, b] = ps.printTitleCols
      if (b >= a && a <= range.left && b < range.right) {
        titleCols = [a, b]
        bodyLeft = Math.max(range.left, b + 1)
      }
    }

    const render = (rg: MergeRange, scale?: number) =>
      r.exportToCanvas({
        range: rg,
        scale: scale ?? opts.scale,
        includeHeaders: opts.includeHeaders,
        gridlines: opts.gridlines,
        background: opts.background,
      })

    const base = render({ top: bodyTop, left: bodyLeft, bottom: range.bottom, right: range.right })
    const deco = await this.collectDecorations(s, base.metrics)
    compositeOverlays(base, deco)
    const S = base.scale // 标题条用同一 scale → 与正文等宽/等高,可逐页拼接

    let repeatTop: ExportSheetImage['repeatTop']
    let repeatLeft: ExportSheetImage['repeatLeft']
    let corner: ExportSheetImage['corner']
    if (titleRows) {
      const strip = render({ top: titleRows[0], bottom: titleRows[1], left: bodyLeft, right: range.right }, S)
      repeatTop = { canvas: strip.canvas, heightCss: strip.bodyH }
    }
    if (titleCols) {
      const strip = render({ top: bodyTop, bottom: range.bottom, left: titleCols[0], right: titleCols[1] }, S)
      repeatLeft = { canvas: strip.canvas, widthCss: strip.bodyW }
    }
    if (titleRows && titleCols) {
      const c = render({ top: titleRows[0], bottom: titleRows[1], left: titleCols[0], right: titleCols[1] }, S)
      corner = { canvas: c.canvas }
    }

    // 打印缩放: 非 fitToWidth 时套用 pageSetup.scale
    const zoom = opts.fitToWidth === false && ps?.scale ? ps.scale / 100 : undefined

    return {
      canvas: base.canvas,
      bodyWcss: base.bodyW,
      bodyHcss: base.bodyH,
      sheetName: s.name,
      repeatTop,
      repeatLeft,
      corner,
      zoom,
    }
  }

  /** 为一个工作表生成矢量导出输入(逐格信息 + 兜底底图 + 图片图表 + 标题行) */
  private async buildVectorSheet(sheetIdx: number, opts: PdfExportOptions): Promise<VectorSheet | null> {
    const wb = this.host.getWorkbook()
    const s = wb?.sheets[sheetIdx]
    if (!wb || !s) return null
    const r = this.rendererFor(sheetIdx, s, wb)
    const metrics = new GridMetrics(s, 1)
    const ps = s.pageSetup
    const dim = s.dimension
    const full: MergeRange = { top: 0, left: 0, bottom: Math.max(0, dim.rows - 1), right: Math.max(0, dim.cols - 1) }
    const range = opts.range ?? ps?.printArea ?? full

    let titleRows: [number, number] | undefined
    let titleCols: [number, number] | undefined
    let bodyTop = range.top
    let bodyLeft = range.left
    if (ps?.printTitleRows) {
      const [a, b] = ps.printTitleRows
      if (b >= a && a <= range.top && b < range.bottom) {
        titleRows = [a, b]
        bodyTop = Math.max(range.top, b + 1)
      }
    }
    if (ps?.printTitleCols) {
      const [a, b] = ps.printTitleCols
      if (b >= a && a <= range.left && b < range.right) {
        titleCols = [a, b]
        bodyLeft = Math.max(range.left, b + 1)
      }
    }
    // 兜底底图须覆盖 标题行/列 ∪ 正文(标题格无字体中文也要能裁图)
    const rasterTop = titleRows ? Math.min(range.top, titleRows[0]) : range.top
    const rasterLeft = titleCols ? Math.min(range.left, titleCols[0]) : range.left

    const base = r.exportToCanvas({
      range: { top: rasterTop, left: rasterLeft, bottom: range.bottom, right: range.right },
      scale: opts.scale,
      includeHeaders: false,
      gridlines: opts.gridlines,
      background: opts.background,
    })
    const deco = await this.collectDecorations(s, base.metrics)
    const images = [...(deco.images ?? []), ...(deco.charts ?? [])]
    const zoom = opts.fitToWidth === false && ps?.scale ? ps.scale / 100 : undefined

    return {
      sheetName: s.name,
      metrics,
      bodyLeft,
      bodyRight: range.right,
      bodyTop,
      bodyBottom: range.bottom,
      titleRows,
      titleCols,
      merges: s.merges,
      gridlines: opts.gridlines ?? s.showGridLines,
      zoom,
      getCell: (rr, cc) => r.exportCellDraw(rr, cc),
      rasterCanvas: base.canvas,
      rasterScale: base.scale,
      rasterTop,
      rasterLeft,
      images,
    }
  }

  /** 从工作表原生 pageSetup 推导导出默认值(纸张/方向/边距/是否适应页宽) */
  private pageSetupDefaults(sheetIdx: number): Partial<PdfExportOptions> {
    const ps = this.host.getWorkbook()?.sheets[sheetIdx]?.pageSetup
    if (!ps) return {}
    const d: Partial<PdfExportOptions> = {}
    if (ps.paperFormat) d.format = ps.paperFormat
    if (ps.orientation) d.orientation = ps.orientation
    if (ps.margins) d.margin = { top: ps.margins.top, right: ps.margins.right, bottom: ps.margins.bottom, left: ps.margins.left }
    // fitToPage → 适应页宽;否则按自然尺寸×scale(buildSheetImage 用 zoom 处理)
    d.fitToWidth = ps.fitToPage ? true : false
    return d
  }

  // ====================== 对外导出 API ======================

  /** 导出当前/指定表为图片 Blob(图片为单表;多表请用 PDF) */
  async exportImage(opts: ImageExportOptions = {}): Promise<Blob> {
    const targets = this.resolveTargets(opts.target)
    if (!targets.length) throw new Error('无可导出的工作表')
    const img = await this.buildSheetImage(targets[0], opts)
    if (!img) throw new Error('导出失败: 无法生成底图')
    return canvasToBlob(img.canvas, opts.type ?? 'png', opts.quality ?? 0.92)
  }
  async downloadImage(opts: ImageExportOptions = {}): Promise<void> {
    const blob = await this.exportImage(opts)
    const ext = opts.type === 'jpeg' ? 'jpg' : opts.type === 'webp' ? 'webp' : 'png'
    downloadBlob(blob, opts.fileName ?? `${this.baseName()}.${ext}`)
  }

  /** 导出为 PDF Blob(每个目标表分页;需可选依赖 jspdf)。未显式指定的页面参数取自工作表 pageSetup。 */
  async exportPdf(opts: PdfExportOptions = {}): Promise<Blob> {
    const targets = this.resolveTargets(opts.target)
    if (!targets.length) throw new Error('无可导出的工作表')
    const eff: PdfExportOptions = { ...this.pageSetupDefaults(targets[0]), ...opts }
    if (eff.vector) {
      const vs = (await Promise.all(targets.map((i) => this.buildVectorSheet(i, eff)))).filter(Boolean) as VectorSheet[]
      return exportToVectorPdf(vs, eff)
    }
    const images = (await Promise.all(targets.map((i) => this.buildSheetImage(i, eff, true)))).filter(Boolean) as ExportSheetImage[]
    return exportToPdf(images, eff)
  }
  async downloadPdf(opts: PdfExportOptions = {}): Promise<void> {
    const blob = await this.exportPdf(opts)
    downloadBlob(blob, opts.fileName ?? `${this.baseName()}.pdf`)
  }

  /** 打开系统打印(可在对话框另存为 PDF)。页面参数同样默认取自 pageSetup。 */
  async print(opts: PrintOptions = {}): Promise<void> {
    const targets = this.resolveTargets(opts.target)
    if (!targets.length) return
    const eff: PrintOptions = { ...this.pageSetupDefaults(targets[0]), ...opts }
    const images = (await Promise.all(targets.map((i) => this.buildSheetImage(i, eff, true)))).filter(Boolean) as ExportSheetImage[]
    printSheets(images, { ...eff, title: eff.title ?? this.baseName() })
  }
}
