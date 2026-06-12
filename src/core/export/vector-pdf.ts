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
  /** 正文列范围(0-based 闭区间,已剔除标题列) */
  bodyLeft: number
  bodyRight: number
  /** 正文行范围(0-based 闭区间,已剔除标题行) */
  bodyTop: number
  bodyBottom: number
  /** 打印标题行 [r0,r1](每页顶部重复) */
  titleRows?: [number, number]
  /** 打印标题列 [c0,c1](每页左侧重复) */
  titleCols?: [number, number]
  merges: MergeRange[]
  gridlines: boolean
  /** 非 fitToWidth 时的打印缩放(pageSetup.scale/100) */
  zoom?: number
  /** 取一个格的绘制信息 */
  getCell: (row: number, col: number) => CellDrawInfo | null
  /** 栅格兜底底图: 覆盖 [rasterTop..bodyBottom] 行 × [rasterLeft..bodyRight] 列,设备像素 */
  rasterCanvas: HTMLCanvasElement
  rasterScale: number
  rasterTop: number
  rasterLeft: number
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
  titleColCols: number[] // 标题列(列号)
  titleRowRows: number[] // 标题行(行号)
  titleColWmm: number // 标题列总宽 mm
  titleRowHmm: number // 标题行总高 mm
  pages: { cols: number[]; rows: number[] }[] // 每页的正文列/行号
}

/** 把一串索引按各自尺寸打包成多段(每段累计 ≤ maxMm;每段至少 1 个)。(导出供测) */
export function packBands(indices: number[], sizeMm: (i: number) => number, maxMm: number): number[][] {
  const out: number[][] = []
  let cur: number[] = []
  let acc = 0
  for (const idx of indices) {
    const sz = sizeMm(idx)
    if (cur.length && acc + sz > maxMm) {
      out.push(cur)
      cur = []
      acc = 0
    }
    cur.push(idx)
    acc += sz
  }
  if (cur.length) out.push(cur)
  if (!out.length) out.push([])
  return out
}

/** 二维分页: 列带 × 行带(先下后右) */
function planSheet(vs: VectorSheet, page: { contentWmm: number; contentHmm: number; fitToWidth: boolean }): SheetPlan {
  const m = vs.metrics
  const bodyWcss = m.colLeft(vs.bodyRight + 1) - m.colLeft(vs.bodyLeft)
  const titleColCols = vs.titleCols ? rangeArr(vs.titleCols[0], vs.titleCols[1]) : []
  const titleRowRows = vs.titleRows ? rangeArr(vs.titleRows[0], vs.titleRows[1]) : []
  const titleColWcss = titleColCols.reduce((s, c) => s + m.colWidth(c), 0)

  const zoom = vs.zoom && vs.zoom > 0 ? vs.zoom : 1
  const factor = page.fitToWidth
    ? (bodyWcss + titleColWcss) * MM_PER_PX > 0
      ? page.contentWmm / ((bodyWcss + titleColWcss) * MM_PER_PX)
      : 1
    : zoom
  const mmPerCss = MM_PER_PX * factor

  const titleColWmm = titleColWcss * mmPerCss
  const titleRowHmm = titleRowRows.reduce((s, r) => s + m.rowHeight(r) * mmPerCss, 0)
  const availWmm = Math.max(1, page.contentWmm - titleColWmm)
  const availHmm = Math.max(1, page.contentHmm - titleRowHmm)

  const bodyCols = rangeArr(vs.bodyLeft, vs.bodyRight)
  const bodyRows = rangeArr(vs.bodyTop, vs.bodyBottom)
  // fitToWidth → 列不分页(整宽一带)
  const colBands = page.fitToWidth ? [bodyCols] : packBands(bodyCols, (c) => m.colWidth(c) * mmPerCss, availWmm)
  const rowBands = packBands(bodyRows, (r) => m.rowHeight(r) * mmPerCss, availHmm)

  const pages: { cols: number[]; rows: number[] }[] = []
  for (const cols of colBands) for (const rows of rowBands) pages.push({ cols, rows })

  return { mmPerCss, titleColCols, titleRowRows, titleColWmm, titleRowHmm, pages }
}

/** 渲染一页: 标题角 + 标题行 + 标题列 + 正文(各区域独立 x/y 映射),末尾贴图片 */
function renderPage(
  doc: any,
  vs: VectorSheet,
  plan: SheetPlan,
  pg: { cols: number[]; rows: number[] },
  margin: { top: number; left: number },
  baseFont: string,
  customFont: boolean,
) {
  const m = vs.metrics
  const { mmPerCss, titleColCols, titleRowRows, titleColWmm, titleRowHmm } = plan
  const bodyX0 = margin.left + titleColWmm
  const bodyY0 = margin.top + titleRowHmm

  // 各区域的列/行 x/y 映射(mm,格子左上角)
  const colXbody = (c: number) => bodyX0 + (m.colLeft(c) - m.colLeft(pg.cols[0] ?? vs.bodyLeft)) * mmPerCss
  const rowYbody = (r: number) => bodyY0 + (m.rowTop(r) - m.rowTop(pg.rows[0] ?? vs.bodyTop)) * mmPerCss
  const colXtitle = (c: number) => margin.left + (m.colLeft(c) - m.colLeft(titleColCols[0] ?? 0)) * mmPerCss
  const rowYtitle = (r: number) => margin.top + (m.rowTop(r) - m.rowTop(titleRowRows[0] ?? 0)) * mmPerCss

  const ctx = { doc, vs, baseFont, customFont, mmPerCss }
  // 正文
  drawRegion(ctx, pg.rows, pg.cols, colXbody, rowYbody)
  // 标题行(× 正文列)
  if (titleRowRows.length) drawRegion(ctx, titleRowRows, pg.cols, colXbody, rowYtitle)
  // 标题列(× 正文行)
  if (titleColCols.length) drawRegion(ctx, pg.rows, titleColCols, colXtitle, rowYbody)
  // 标题角
  if (titleRowRows.length && titleColCols.length) drawRegion(ctx, titleRowRows, titleColCols, colXtitle, rowYtitle)

  // 图片/图表: from.row/col 落在本页正文区才贴
  const rowSet = new Set(pg.rows)
  for (const im of vs.images) {
    const fr = im.anchor.from.row
    const fc = im.anchor.from.col
    if (!rowSet.has(fr) || fc < (pg.cols[0] ?? 0) || fc > (pg.cols[pg.cols.length - 1] ?? -1)) continue
    const rect = anchorRect(m, im.anchor)
    const ix = bodyX0 + (rect.left - m.colLeft(pg.cols[0])) * mmPerCss
    const iy = bodyY0 + (rect.top - m.rowTop(pg.rows[0])) * mmPerCss
    try {
      const url = canvasSourceToDataUrl(im.source)
      if (url) doc.addImage(url, 'PNG', ix, iy, rect.width * mmPerCss, rect.height * mmPerCss, undefined, 'FAST')
    } catch {
      /* 跳过 */
    }
  }
}

interface RegionCtx {
  doc: any
  vs: VectorSheet
  baseFont: string
  customFont: boolean
  mmPerCss: number
}

/** 画一个区域(行集×列集): 逐格矢量/栅格 + 浅网格线;尊重合并 */
function drawRegion(
  ctx: RegionCtx,
  rows: number[],
  cols: number[],
  xOf: (c: number) => number,
  yOf: (r: number) => number,
) {
  const { doc, vs, baseFont, customFont, mmPerCss } = ctx
  if (!rows.length || !cols.length) return
  const m = vs.metrics
  const covered = mergeCoverSet(vs.merges)
  const anchorOf = mergeAnchorMap(vs.merges)
  const cMin = cols[0]
  const cMax = cols[cols.length - 1]

  if (vs.gridlines) drawRegionGrid(doc, m, rows, cols, xOf, yOf, mmPerCss)

  void cMin
  for (const r of rows) {
    const yTop = yOf(r)
    for (let ci = 0; ci < cols.length; ci++) {
      const c = cols[ci]
      const key = r + ':' + c
      if (covered.has(key) && !anchorOf.has(key)) continue
      const mg = anchorOf.get(key)
      const spanRight = mg ? Math.min(mg.right, cMax) : c
      const spanBottom = mg ? mg.bottom : r
      const x = xOf(c)
      const w = (m.colLeft(spanRight + 1) - m.colLeft(c)) * mmPerCss
      const h = (m.rowTop(spanBottom + 1) - m.rowTop(r)) * mmPerCss
      const info = vs.getCell(r, c)
      if (info) {
        const fallback = info.complex || (!customFont && hasNonLatin(info.text))
        if (fallback) rasterCell(doc, vs, r, c, spanRight, spanBottom, x, yTop, w, h)
        else drawVectorCell(doc, info, x, yTop, w, h, mmPerCss, baseFont, customFont)
      }
      if (mg && mg.right > c) {
        // 跳到合并块末列(限本区域内)
        while (ci < cols.length - 1 && cols[ci + 1] <= mg.right) ci++
      }
    }
  }
}

function rangeArr(a: number, b: number): number[] {
  const out: number[] = []
  for (let i = a; i <= b; i++) out.push(i)
  return out
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
  // 条件格式背景(盖在普通填充之上)
  const eff = info.effect
  if (eff?.fillColor) {
    const [r, g, b] = hexToRgb(eff.fillColor)
    doc.setFillColor(r, g, b)
    doc.rect(x, y, w, h, 'F')
  }
  // 数据条
  if (eff?.dataBar) {
    const barW = Math.max(0, (w - 2 * mmPerCss) * eff.dataBar.ratio)
    if (barW > 0) {
      const [r, g, b] = hexToRgb(eff.dataBar.color)
      doc.setFillColor(r, g, b)
      doc.rect(x + mmPerCss, y + mmPerCss, barW, h - 2 * mmPerCss, 'F')
    }
  }
  // 边框
  drawBorder(doc, s.borders.top, x, y, x + w, y)
  drawBorder(doc, s.borders.bottom, x, y + h, x + w, y + h)
  drawBorder(doc, s.borders.left, x, y, x, y + h)
  drawBorder(doc, s.borders.right, x + w, y, x + w, y + h)
  // 对角线边框(↘ / ↗)
  if (s.borders.diagonal && (s.borders.diagonalDown || s.borders.diagonalUp)) {
    if (s.borders.diagonalDown) drawBorder(doc, s.borders.diagonal, x, y, x + w, y + h)
    if (s.borders.diagonalUp) drawBorder(doc, s.borders.diagonal, x, y + h, x + w, y)
  }
  // 条件格式图标(画在左侧,文本相应右移)
  let iconShift = 0
  if (eff?.icon) {
    drawVectorIcon(doc, eff.icon, x + 1.5 * mmPerCss, y + h / 2, Math.min(h * 0.32, 2.2 * mmPerCss))
    iconShift = 5 * mmPerCss
  }
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
    tx = x + pad + iconShift + (s.indent || 0) * 8 * mmPerCss
  }
  let ty: number
  let baseline: 'top' | 'middle' | 'bottom'
  // 溢出顶对齐:wrap 文本折行后总高超过格高 → 顶对齐显示文头(跟 canvas/Excel 一致,避免 PDF 里被裁掉文头)
  let forceTop = false
  if (s.wrapText) {
    const fsMm = s.font.size * (mmPerCss / MM_PER_PX)
    try {
      const lines = doc.splitTextToSize(text, Math.max(1, w - 2 * pad))
      if (lines.length * fsMm * 1.15 > h - 2 * pad) forceTop = true
    } catch { /* splitTextToSize 失败忽略 */ }
  }
  if (s.vAlign === 'top' || forceTop) {
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

/** 条件格式图标(矢量近似): 红黄绿圆点 / 三向箭头 / 色阶圆点 */
function drawVectorIcon(doc: any, icon: { setName: string; level: number; count: number }, cx: number, cy: number, r: number) {
  const name = icon.setName || ''
  if (name.includes('Arrows')) {
    // 箭头: 低→下(红) 中→平(黄) 高→上(绿)
    const colors: [number, number, number][] = [
      [214, 59, 59],
      [232, 181, 58],
      [91, 159, 78],
    ]
    const idx = icon.level === 0 ? 0 : icon.level >= icon.count - 1 ? 2 : 1
    const [cr, cg, cb] = colors[idx]
    doc.setDrawColor(cr, cg, cb)
    doc.setLineWidth(r * 0.35)
    if (idx === 1) {
      doc.line(cx - r, cy, cx + r, cy) // 平
    } else {
      const dy = idx === 0 ? r : -r // 下 / 上
      doc.line(cx, cy - dy, cx, cy + dy)
      doc.line(cx, cy + dy, cx - r * 0.6, cy + dy - Math.sign(dy) * r * 0.6)
      doc.line(cx, cy + dy, cx + r * 0.6, cy + dy - Math.sign(dy) * r * 0.6)
    }
    return
  }
  // 默认: 红/黄/绿圆点(交通灯/符号),或按色阶
  const palette: [number, number, number][] = [
    [214, 59, 59],
    [232, 181, 58],
    [91, 159, 78],
  ]
  const t = icon.count > 1 ? icon.level / (icon.count - 1) : 1
  let col: [number, number, number]
  if (name.includes('TrafficLights') || name.includes('Signs') || name.includes('Symbols')) {
    col = palette[Math.min(icon.level, 2)]
  } else {
    col = [Math.round(214 - 123 * t), Math.round(59 + 100 * t), 59]
  }
  doc.setFillColor(col[0], col[1], col[2])
  doc.circle(cx, cy, r, 'F')
}

function drawBorder(doc: any, edge: any, x1: number, y1: number, x2: number, y2: number) {
  if (!edge || !edge.style || edge.style === 'none') return
  const [r, g, b] = hexToRgb(edge.color || '#000000')
  doc.setDrawColor(r, g, b)
  const lw = BORDER_W[edge.style] ?? 0.2
  if (edge.style === 'double') {
    // double 画成沿法线偏移的两条平行线(跟 canvas 一致,不再退化成单线)
    doc.setLineWidth(Math.max(0.1, lw * 0.5))
    const dx = x2 - x1, dy = y2 - y1, len = Math.hypot(dx, dy) || 1
    const off = Math.max(0.18, lw * 0.6)
    const nx = (-dy / len) * off, ny = (dx / len) * off
    doc.line(x1 + nx, y1 + ny, x2 + nx, y2 + ny)
    doc.line(x1 - nx, y1 - ny, x2 - nx, y2 - ny)
    return
  }
  doc.setLineWidth(lw)
  doc.line(x1, y1, x2, y2)
}

/** 浅灰网格线: 给一个区域画行/列边线(线很浅,边框会覆盖在其上) */
function drawRegionGrid(
  doc: any,
  m: GridMetrics,
  rows: number[],
  cols: number[],
  xOf: (c: number) => number,
  yOf: (r: number) => number,
  mmPerCss: number,
) {
  doc.setDrawColor(224, 226, 229)
  doc.setLineWidth(0.1)
  const cFirst = cols[0]
  const cLast = cols[cols.length - 1]
  const rFirst = rows[0]
  const rLast = rows[rows.length - 1]
  const x0 = xOf(cFirst)
  const x1 = xOf(cLast) + m.colWidth(cLast) * mmPerCss
  const y0 = yOf(rFirst)
  const y1 = yOf(rLast) + m.rowHeight(rLast) * mmPerCss
  for (const r of rows) doc.line(x0, yOf(r), x1, yOf(r)) // 行顶线
  doc.line(x0, y1, x1, y1) // 末行底线
  for (const c of cols) doc.line(xOf(c), y0, xOf(c), y1) // 列左线
  doc.line(x1, y0, x1, y1) // 末列右线
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
  const sx = (m.colLeft(c) - m.colLeft(vs.rasterLeft)) * vs.rasterScale
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
