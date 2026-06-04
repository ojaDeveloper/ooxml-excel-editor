/**
 * 矢量 PDF: 逐格用 jsPDF 真文字 + 矢量填充/边框绘制(文字可选可搜、清晰、文件小)。
 * 与位图 PDF(pdf.ts)并存。难啃的格(条件格式效果/迷你图/旋转/富文本,或无中文字体时的非拉丁文本)
 * 从整表底图栅格裁该格小图贴上 —— 内容不丢。
 *
 * 字体: jsPDF 内置字体只认拉丁/数字。宿主用 configureDoc(doc) 钩子 addFont 注册中文字体后,
 * 检测到自定义字体即全矢量;否则非拉丁单元格走栅格兜底。
 */
import type { GridMetrics } from '../layout/grid-metrics'
import type { CellDrawInfo } from '../render/canvas-renderer'
import type { ImageAnchor, MergeRange } from '../model/types'
import type { BeforeRenderPage, PdfExportOptions, PdfPageContext } from './types'
import { MM_PER_PX, resolveMargins, resolvePageSize } from './raster'
import { anchorRect } from '../overlay/anchor'

/** 一个待矢量导出的工作表 */
export interface VectorSheet {
  sheetName: string
  metrics: GridMetrics
  /** 列范围(0-based 闭区间) */
  left: number
  right: number
  /** 正文行范围(0-based 闭区间,已剔除标题行) */
  bodyTop: number
  bodyBottom: number
  /** 打印标题行 [r0,r1](每页顶部重复) */
  titleRows?: [number, number]
  merges: MergeRange[]
  gridlines: boolean
  /** 非 fitToWidth 时的打印缩放(pageSetup.scale/100) */
  zoom?: number
  /** 取一个格的绘制信息 */
  getCell: (row: number, col: number) => CellDrawInfo | null
  /** 栅格兜底底图: 覆盖 [rasterTop..bodyBottom] 行 × [left..right] 列,设备像素 */
  rasterCanvas: HTMLCanvasElement
  rasterScale: number
  rasterTop: number
  /** 图片/图表(已加载源 + 锚点),addImage 贴上 */
  images: { source: CanvasImageSource; anchor: ImageAnchor }[]
}

async function loadJsPdf(): Promise<any> {
  try {
    // @ts-ignore 可选依赖
    const mod: any = await import(/* @vite-ignore */ 'jspdf')
    return mod.jsPDF || mod.default?.jsPDF || mod.default || mod
  } catch (e) {
    throw new Error('矢量 PDF 导出需要可选依赖 jspdf,请先安装: npm i jspdf (' + (e as Error).message + ')')
  }
}

const STD_FONTS = new Set(['helvetica', 'times', 'courier', 'symbol', 'zapfdingbats'])
/** 含非拉丁字符(粗略: 码点 > 0x250) → jsPDF 内置字体画不了,需栅格兜底。(导出供测) */
export function hasNonLatin(s: string): boolean {
  for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) > 0x250) return true
  return false
}
/** "#RRGGBB" → [r,g,b];非法返回黑。(导出供测) */
export function hexToRgb(hex: string): [number, number, number] {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex)
  if (!m) return [0, 0, 0]
  const n = parseInt(m[1], 16)
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255]
}
const BORDER_W: Record<string, number> = { thin: 0.2, hair: 0.1, medium: 0.45, thick: 0.7, double: 0.45 }

export async function exportToVectorPdf(sheets: VectorSheet[], opts: PdfExportOptions = {}): Promise<Blob> {
  if (!sheets.length) throw new Error('没有可导出的工作表')
  const JsPDF = await loadJsPdf()
  const [pageW, pageH] = resolvePageSize(opts.format, opts.orientation)
  const margin = resolveMargins(opts.margin)
  const contentWmm = Math.max(1, pageW - margin.left - margin.right)
  const contentHmm = Math.max(1, pageH - margin.top - margin.bottom)
  const fitToWidth = opts.fitToWidth ?? true

  const doc = new JsPDF({ unit: 'mm', format: [pageW, pageH], orientation: 'portrait' })
  // 字体注册钩子(宿主可 addFont 中文字体)
  if (opts.configureDoc) {
    try {
      opts.configureDoc(doc)
    } catch (e) {
      console.warn('[ooxml-preview] configureDoc 抛错:', e)
    }
  }
  const activeFont: string = (doc.getFont?.() || {}).fontName || 'helvetica'
  const customFont = !STD_FONTS.has(String(activeFont).toLowerCase())
  // 自定义字体设为默认(宿主已 setFont 时尊重其选择)
  const baseFont = customFont ? activeFont : 'helvetica'

  // 先把每个表分页,算总页数(给钩子)
  const plans = sheets.map((vs) => planSheet(vs, { contentWmm, contentHmm, fitToWidth }))
  const pageCount = plans.reduce((n, p) => n + p.pages.length, 0) || 1

  let globalPage = 0
  let first = true
  for (let si = 0; si < sheets.length; si++) {
    const vs = sheets[si]
    const plan = plans[si]
    for (const pg of plan.pages) {
      if (!first) doc.addPage([pageW, pageH], 'portrait')
      first = false
      renderPage(doc, vs, plan, pg, margin, baseFont, customFont)
      runBeforeRenderPage(opts.beforeRenderPage, {
        doc,
        pageIndex: globalPage,
        pageCount,
        sheetName: vs.sheetName,
        sheetIndex: si,
        pageWidth: pageW,
        pageHeight: pageH,
        margin,
      })
      globalPage++
    }
  }
  return doc.output('blob')
}

function runBeforeRenderPage(fn: BeforeRenderPage | undefined, ctx: PdfPageContext) {
  if (!fn) return
  try {
    fn(ctx)
  } catch (e) {
    console.warn('[ooxml-preview] beforeRenderPage 抛错:', e)
  }
}

interface SheetPlan {
  mmPerCss: number
  ox: number // range 左上角 css(列)
  oy: number
  titleRows: number[]
  titleHmm: number
  pages: { rows: number[] }[] // 每页的正文行号列表
}

/** 分页: 按行高把正文行打包到各页(每页至少一行) */
function planSheet(vs: VectorSheet, page: { contentWmm: number; contentHmm: number; fitToWidth: boolean }): SheetPlan {
  const m = vs.metrics
  const ox = m.colLeft(vs.left)
  const oy = m.rowTop(vs.bodyTop)
  const bodyWcss = m.colLeft(vs.right + 1) - ox
  const naturalWmm = bodyWcss * MM_PER_PX
  const zoom = vs.zoom && vs.zoom > 0 ? vs.zoom : 1
  const drawWmm = page.fitToWidth ? page.contentWmm : Math.min(naturalWmm * zoom, page.contentWmm)
  const factor = naturalWmm > 0 ? drawWmm / naturalWmm : 1
  const mmPerCss = MM_PER_PX * factor

  const titleRows: number[] = []
  if (vs.titleRows) {
    for (let r = vs.titleRows[0]; r <= vs.titleRows[1]; r++) titleRows.push(r)
  }
  const titleHmm = titleRows.reduce((s, r) => s + m.rowHeight(r) * mmPerCss, 0)
  const availHmm = Math.max(1, page.contentHmm - titleHmm)

  const pages: { rows: number[] }[] = []
  let cur: number[] = []
  let acc = 0
  for (let r = vs.bodyTop; r <= vs.bodyBottom; r++) {
    const hmm = m.rowHeight(r) * mmPerCss
    if (cur.length && acc + hmm > availHmm) {
      pages.push({ rows: cur })
      cur = []
      acc = 0
    }
    cur.push(r)
    acc += hmm
  }
  if (cur.length) pages.push({ rows: cur })
  if (!pages.length) pages.push({ rows: [] })

  return { mmPerCss, ox, oy: oy, titleRows, titleHmm, pages }
}

/** 渲染一页: 标题行(若有) + 正文行,末尾贴图片/图表 */
function renderPage(
  doc: any,
  vs: VectorSheet,
  plan: SheetPlan,
  pg: { rows: number[] },
  margin: { top: number; left: number },
  baseFont: string,
  customFont: boolean,
) {
  const m = vs.metrics
  const { mmPerCss, ox } = plan
  const colX = (c: number) => margin.left + (m.colLeft(c) - ox) * mmPerCss

  // 行号 → 本页 y 起点(mm)。先标题行,再正文行。
  const rowY = new Map<number, number>()
  let y = margin.top
  for (const r of plan.titleRows) {
    rowY.set(r, y)
    y += m.rowHeight(r) * mmPerCss
  }
  const bodyStartY = y
  for (const r of pg.rows) {
    rowY.set(r, y)
    y += m.rowHeight(r) * mmPerCss
  }

  const drawnRows = [...plan.titleRows, ...pg.rows]
  const covered = mergeCoverSet(vs.merges)
  const anchorOf = mergeAnchorMap(vs.merges)

  for (const r of drawnRows) {
    const yTop = rowY.get(r)!
    for (let c = vs.left; c <= vs.right; c++) {
      const key = r + ':' + c
      if (covered.has(key) && !anchorOf.has(key)) continue
      const mg = anchorOf.get(key)
      const spanRight = mg ? mg.right : c
      const spanBottom = mg ? mg.bottom : r
      const x = colX(c)
      const w = colX(spanRight + 1) - x
      const h = (m.rowTop(spanBottom + 1) - m.rowTop(r)) * mmPerCss

      const info = vs.getCell(r, c)
      if (!info) continue
      const fallback = info.complex || (!customFont && hasNonLatin(info.text))
      if (fallback) {
        rasterCell(doc, vs, r, c, spanRight, spanBottom, x, yTop, w, h)
      } else {
        drawVectorCell(doc, info, x, yTop, w, h, mmPerCss, baseFont, customFont)
      }
      if (mg && mg.right > c) c = mg.right // 跳过被合并覆盖的后续列
    }
  }

  // 轻网格线(可选;边框已在格上画,这里补无边框格的浅灰线)
  if (vs.gridlines) drawGridlines(doc, vs, plan, drawnRows, rowY, margin)

  // 图片/图表: 贴到其 from.row 所在页(本页)
  for (const im of vs.images) {
    const fr = im.anchor.from.row
    if (!rowY.has(fr)) continue
    const rect = anchorRect(m, im.anchor)
    const ix = margin.left + (rect.left - ox) * mmPerCss
    const iyRow = rowY.get(fr)!
    const iy = iyRow + (rect.top - m.rowTop(fr)) * mmPerCss
    try {
      const url = canvasSourceToDataUrl(im.source)
      if (url) doc.addImage(url, 'PNG', ix, iy, rect.width * mmPerCss, rect.height * mmPerCss, undefined, 'FAST')
    } catch {
      /* 跳过 */
    }
  }
  void bodyStartY
}

function drawVectorCell(
  doc: any,
  info: CellDrawInfo,
  x: number,
  y: number,
  w: number,
  h: number,
  mmPerCss: number,
  baseFont: string,
  customFont: boolean,
) {
  const s = info.style
  // 填充
  if (s.fill.type === 'solid' && s.fill.fgColor) {
    const [r, g, b] = hexToRgb(s.fill.fgColor)
    doc.setFillColor(r, g, b)
    doc.rect(x, y, w, h, 'F')
  }
  // 边框
  drawBorder(doc, s.borders.top, x, y, x + w, y)
  drawBorder(doc, s.borders.bottom, x, y + h, x + w, y + h)
  drawBorder(doc, s.borders.left, x, y, x, y + h)
  drawBorder(doc, s.borders.right, x + w, y, x + w, y + h)
  // 文本
  const text = info.text
  if (!text) return
  const pad = 2 * mmPerCss
  // 自定义字体通常只注册了 normal 字重,请求未注册的 bold 会抛错 → 仅内置字体用 bold
  try {
    doc.setFont(baseFont, !customFont && info.bold ? 'bold' : 'normal')
  } catch {
    doc.setFont(baseFont, 'normal')
  }
  doc.setFontSize(s.font.size * (mmPerCss / MM_PER_PX))
  const [tr, tg, tb] = hexToRgb(info.color)
  doc.setTextColor(tr, tg, tb)

  const isNumber = /^[\s\d.,%$()+\-eE/]+$/.test(text)
  const hAlign = s.hAlign === 'general' ? (isNumber ? 'right' : 'left') : s.hAlign
  let tx: number
  let align: 'left' | 'center' | 'right' = 'left'
  if (hAlign === 'center') {
    tx = x + w / 2
    align = 'center'
  } else if (hAlign === 'right') {
    tx = x + w - pad
    align = 'right'
  } else {
    tx = x + pad + (s.indent || 0) * 8 * mmPerCss
  }
  let ty: number
  let baseline: 'top' | 'middle' | 'bottom'
  if (s.vAlign === 'top') {
    ty = y + pad
    baseline = 'top'
  } else if (s.vAlign === 'middle') {
    ty = y + h / 2
    baseline = 'middle'
  } else {
    ty = y + h - pad
    baseline = 'bottom'
  }
  const tOpts: any = { align, baseline }
  if (s.wrapText) tOpts.maxWidth = Math.max(1, w - 2 * pad)
  try {
    doc.text(text, tx, ty, tOpts)
  } catch {
    /* 个别字符画不出时跳过该格文本(不影响其它) */
  }
}

function drawBorder(doc: any, edge: any, x1: number, y1: number, x2: number, y2: number) {
  if (!edge || !edge.style || edge.style === 'none') return
  const [r, g, b] = hexToRgb(edge.color || '#000000')
  doc.setDrawColor(r, g, b)
  doc.setLineWidth(BORDER_W[edge.style] ?? 0.2)
  doc.line(x1, y1, x2, y2)
}

/** 浅灰网格线: 给整片正文画 cell 边线(在内容之下视觉上够用,这里画在最后,线很浅) */
function drawGridlines(
  doc: any,
  vs: VectorSheet,
  plan: SheetPlan,
  rows: number[],
  rowY: Map<number, number>,
  margin: { left: number },
) {
  const m = vs.metrics
  const { mmPerCss, ox } = plan
  doc.setDrawColor(224, 226, 229)
  doc.setLineWidth(0.1)
  for (const r of rows) {
    const y = rowY.get(r)!
    const h = m.rowHeight(r) * mmPerCss
    const x0 = margin.left + (m.colLeft(vs.left) - ox) * mmPerCss
    const x1 = margin.left + (m.colLeft(vs.right + 1) - ox) * mmPerCss
    doc.line(x0, y + h, x1, y + h) // 行底线
  }
  for (let c = vs.left; c <= vs.right + 1; c++) {
    const x = margin.left + (m.colLeft(c) - ox) * mmPerCss
    const yTop = rowY.get(rows[0])!
    const lastR = rows[rows.length - 1]
    const yBot = rowY.get(lastR)! + m.rowHeight(lastR) * mmPerCss
    doc.line(x, yTop, x, yBot)
  }
}

/** 把某格从底图栅格裁出贴上(兜底) */
function rasterCell(
  doc: any,
  vs: VectorSheet,
  r: number,
  c: number,
  spanRight: number,
  spanBottom: number,
  x: number,
  y: number,
  w: number,
  h: number,
) {
  const m = vs.metrics
  const sx = (m.colLeft(c) - m.colLeft(vs.left)) * vs.rasterScale
  const sy = (m.rowTop(r) - m.rowTop(vs.rasterTop)) * vs.rasterScale
  const sw = (m.colLeft(spanRight + 1) - m.colLeft(c)) * vs.rasterScale
  const sh = (m.rowTop(spanBottom + 1) - m.rowTop(r)) * vs.rasterScale
  if (sw < 1 || sh < 1) return
  const tmp = document.createElement('canvas')
  tmp.width = Math.round(sw)
  tmp.height = Math.round(sh)
  const ctx = tmp.getContext('2d')
  if (!ctx) return
  try {
    ctx.drawImage(vs.rasterCanvas, Math.round(sx), Math.round(sy), tmp.width, tmp.height, 0, 0, tmp.width, tmp.height)
    doc.addImage(tmp.toDataURL('image/png'), 'PNG', x, y, w, h, undefined, 'FAST')
  } catch {
    /* 跳过 */
  }
}

function canvasSourceToDataUrl(src: CanvasImageSource): string | null {
  if (typeof HTMLCanvasElement !== 'undefined' && src instanceof HTMLCanvasElement) return src.toDataURL('image/png')
  if (typeof HTMLImageElement !== 'undefined' && src instanceof HTMLImageElement) {
    const c = document.createElement('canvas')
    c.width = src.naturalWidth || src.width
    c.height = src.naturalHeight || src.height
    const ctx = c.getContext('2d')
    if (!ctx) return null
    try {
      ctx.drawImage(src, 0, 0)
      return c.toDataURL('image/png')
    } catch {
      return null
    }
  }
  return null
}

function mergeCoverSet(merges: MergeRange[]): Set<string> {
  const set = new Set<string>()
  for (const mg of merges) {
    for (let r = mg.top; r <= mg.bottom; r++) for (let c = mg.left; c <= mg.right; c++) set.add(r + ':' + c)
  }
  return set
}
function mergeAnchorMap(merges: MergeRange[]): Map<string, MergeRange> {
  const map = new Map<string, MergeRange>()
  for (const mg of merges) map.set(mg.top + ':' + mg.left, mg)
  return map
}
