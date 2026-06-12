/**
 * ExcelJS Workbook → 中间模型(SheetModel[])。
 * 负责: 单元格值/类型、样式(去重成 styles 数组)、合并、列宽行高、冻结、
 * 条件格式、自动筛选。图片/图表由 drawing-parser/chart-parser 另外补齐。
 */
import type ExcelJS from 'exceljs'
import type {
  CellModel,
  CellStyle,
  ConditionalRule,
  Fill,
  Font,
  MergeRange,
  PageSetupModel,
  RichTextRun,
  SheetModel,
  CssColor,
} from '../model/types'
import { cellKey, makeDefaultStyle } from '../model/types'
import type { ProgressFn } from '../progress'
import { resolveColor } from '../format/color'
import { colWidthToPx, rowHeightToPx, DEFAULT_COL_WIDTH_CHARS, DEFAULT_ROW_HEIGHT_PT } from '../layout/units'

const DEFAULT_FONT: Font = {
  name: 'Calibri',
  size: 11,
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  color: '#000000',
}

export function buildSheets(wb: ExcelJS.Workbook, themeColors: CssColor[], onProgress?: ProgressFn): SheetModel[] {
  const totalRows = wb.worksheets.reduce((s, ws) => s + (ws.rowCount || 0), 0) || 1
  let doneRows = 0
  const onRow = onProgress
    ? () => {
        doneRows++
        if (doneRows % 1000 === 0) onProgress({ stage: 'build', ratio: Math.min(1, doneRows / totalRows) })
      }
    : undefined

  const sheets: SheetModel[] = []
  wb.worksheets.forEach((ws, idx) => {
    sheets.push(buildSheet(ws, idx, themeColors, onRow))
  })
  onProgress?.({ stage: 'build', ratio: 1 })
  return sheets
}

function buildSheet(ws: ExcelJS.Worksheet, index: number, theme: CssColor[], onRow?: () => void): SheetModel {
  // styles[0] 必须是中性空白默认(见 makeDefaultStyle 注释)。绝不能让"第一个被解析到的格样式"占据 index 0,
  // 否则首格(常是带色表头,如 A1 绿底)会成为所有空格/新建格的默认底色,造成粘贴/编辑串色。
  const styles: CellStyle[] = [makeDefaultStyle()]
  const styleIndex = new Map<string, number>([[JSON.stringify(styles[0]), 0]])
  const cells = new Map<string, CellModel>()

  const internStyle = (style: CellStyle): number => {
    const key = JSON.stringify(style)
    const existing = styleIndex.get(key)
    if (existing !== undefined) return existing
    const id = styles.length
    styles.push(style)
    styleIndex.set(key, id)
    return id
  }

  let maxRow = 0
  let maxCol = 0

  // includeEmpty:true —— 必须连**空但带样式**的格也遍历(只有边框/填充的结构格),否则它们的边框/底色会丢
  // (ExcelJS includeEmpty:false 跳过 value 为空的格)。toCellModel 对"空且无可见样式(无边框/填充)"的格返 null,
  // 不入模型,避免把真正空白格也塞进来。eachCell 只到该行最右有格的列,不会扫到无限远。
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const r = rowNumber - 1
      const c = colNumber - 1
      const model = toCellModel(cell, r, c, theme, internStyle)
      if (model) {
        cells.set(cellKey(r, c), model)
        if (r > maxRow) maxRow = r
        if (c > maxCol) maxCol = c
      }
    })
    onRow?.()
  })

  // 合并
  const merges = parseMerges(ws)
  for (const m of merges) {
    maxRow = Math.max(maxRow, m.bottom)
    maxCol = Math.max(maxCol, m.right)
  }

  // 列宽 / 行高
  const columns = new Map<number, { width: number; hidden: boolean }>()
  const props: any = (ws as any).properties || {}
  const defaultColWidth = colWidthToPx(props.defaultColWidth ?? DEFAULT_COL_WIDTH_CHARS)
  const defaultRowHeight = rowHeightToPx(props.defaultRowHeight ?? DEFAULT_ROW_HEIGHT_PT)

  for (let i = 1; i <= ws.columnCount; i++) {
    const col = ws.getColumn(i)
    if (col && (col.width != null || col.hidden)) {
      columns.set(i - 1, {
        width: col.width != null ? colWidthToPx(col.width) : defaultColWidth,
        hidden: !!col.hidden,
      })
    }
  }

  const rows = new Map<number, { height: number; hidden: boolean }>()
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (row.height != null || row.hidden) {
      rows.set(rowNumber - 1, {
        height: row.height != null ? rowHeightToPx(row.height) : defaultRowHeight,
        hidden: !!row.hidden,
      })
    }
  })

  // 冻结
  const view = (ws.views && ws.views[0]) as any
  const freeze = {
    frozenRows: view?.state === 'frozen' ? view.ySplit ?? 0 : 0,
    frozenCols: view?.state === 'frozen' ? view.xSplit ?? 0 : 0,
  }
  const showGridLines = view?.showGridLines !== false

  // 条件格式
  const conditional = parseConditional(ws, theme)

  // 自动筛选
  const autoFilterRange = parseAutoFilter(ws)

  // 数据验证(只取 list 型，用于画下拉)
  const dataValidations = parseDataValidations(ws)

  // 原生页面设置(导出/打印默认值)
  const pageSetup = parsePageSetup(ws)

  const state = (ws as any).state === 'hidden'
    ? 'hidden'
    : (ws as any).state === 'veryHidden'
      ? 'veryHidden'
      : 'visible'

  return {
    name: ws.name,
    index,
    state,
    dimension: { rows: maxRow + 1, cols: maxCol + 1 },
    cells,
    styles,
    merges,
    columns,
    rows,
    defaultColWidth,
    defaultRowHeight,
    freeze,
    conditional,
    autoFilterRange,
    dataValidations,
    images: [],
    charts: [],
    shapes: [],
    sparklines: [],
    pivotTables: [],
    pageSetup,
    showGridLines,
  }
}

/** paperSize 代码 → PageFormat(只映射常见几种,其余省略走默认 a4) */
const PAPER_SIZE_MAP: Record<number, PageSetupModel['paperFormat']> = {
  1: 'letter', // Letter 8.5x11
  8: 'a3', // A3 297x420
  9: 'a4', // A4 210x297
  70: 'a4', // ISO A4 (部分写法)
}
const IN_TO_MM = 25.4

/** 从 ExcelJS worksheet.pageSetup 抽出原生页面设置;失败返回 undefined。(导出供单测) */
export function parsePageSetup(ws: ExcelJS.Worksheet): PageSetupModel | undefined {
  const ps: any = (ws as any).pageSetup
  if (!ps) return undefined
  const out: PageSetupModel = {}

  if (ps.orientation === 'portrait' || ps.orientation === 'landscape') out.orientation = ps.orientation
  if (typeof ps.paperSize === 'number' && PAPER_SIZE_MAP[ps.paperSize]) out.paperFormat = PAPER_SIZE_MAP[ps.paperSize]
  if (typeof ps.scale === 'number' && ps.scale > 0) out.scale = ps.scale
  if (ps.fitToPage) out.fitToPage = true
  if (typeof ps.fitToWidth === 'number') out.fitToWidth = ps.fitToWidth
  if (typeof ps.fitToHeight === 'number') out.fitToHeight = ps.fitToHeight

  const m = ps.margins
  if (m && typeof m === 'object') {
    out.margins = {
      top: (m.top ?? 0.75) * IN_TO_MM,
      bottom: (m.bottom ?? 0.75) * IN_TO_MM,
      left: (m.left ?? 0.7) * IN_TO_MM,
      right: (m.right ?? 0.7) * IN_TO_MM,
      header: (m.header ?? 0.3) * IN_TO_MM,
      footer: (m.footer ?? 0.3) * IN_TO_MM,
    }
  }

  // 打印区域(可能多段,取第一段)
  if (typeof ps.printArea === 'string' && ps.printArea.trim()) {
    const first = ps.printArea.split(',')[0]
    const rg = parseA1Range(stripSheetRef(first))
    if (rg) out.printArea = rg
  }

  // 打印标题行/列: ExcelJS 给 "1:3" / "A:B"(可能带 Sheet! 前缀)
  const rows = parseRowSpan(ps.printTitlesRow)
  if (rows) out.printTitleRows = rows
  const cols = parseColSpan(ps.printTitlesColumn)
  if (cols) out.printTitleCols = cols

  return Object.keys(out).length ? out : undefined
}

/** 去掉 "Sheet1!$A$1:$B$2" 里的 "Sheet1!" 前缀 */
function stripSheetRef(ref: string): string {
  const i = ref.lastIndexOf('!')
  return (i >= 0 ? ref.slice(i + 1) : ref).trim()
}

/** "1:3" / "$1:$3" / "Sheet!$1:$3" → [0,2] 0-based 行区间 */
function parseRowSpan(spec: any): [number, number] | undefined {
  if (typeof spec !== 'string' || !spec.trim()) return undefined
  const s = stripSheetRef(spec).replace(/\$/g, '')
  const m = /^(\d+):(\d+)$/.exec(s)
  if (!m) return undefined
  const a = parseInt(m[1], 10) - 1
  const b = parseInt(m[2], 10) - 1
  if (!Number.isFinite(a) || !Number.isFinite(b)) return undefined
  return [Math.min(a, b), Math.max(a, b)]
}

/** "A:B" / "$A:$B" / "Sheet!$A:$B" → [0,1] 0-based 列区间 */
function parseColSpan(spec: any): [number, number] | undefined {
  if (typeof spec !== 'string' || !spec.trim()) return undefined
  const s = stripSheetRef(spec).replace(/\$/g, '')
  const m = /^([A-Z]+):([A-Z]+)$/i.exec(s)
  if (!m) return undefined
  const a = colLettersToIndex(m[1].toUpperCase())
  const b = colLettersToIndex(m[2].toUpperCase())
  return [Math.min(a, b), Math.max(a, b)]
}

function parseDataValidations(ws: ExcelJS.Worksheet): MergeRange[] {
  const model: any = (ws as any).dataValidations?.model
  if (!model) return []
  const out: MergeRange[] = []
  for (const [addr, rule] of Object.entries(model)) {
    if ((rule as any)?.type !== 'list') continue
    for (const part of addr.split(/\s+/)) {
      const rg = parseA1Range(part)
      if (rg) out.push(rg)
    }
  }
  return out
}

function toCellModel(
  cell: ExcelJS.Cell,
  row: number,
  col: number,
  theme: CssColor[],
  internStyle: (s: CellStyle) => number,
): CellModel | null {
  const value = cell.value
  if (value === null || value === undefined) {
    // 空格只有"带可见样式(边框/填充)"才入模型 —— 结构格的边框/底色靠这条保住;真正空白格(无边框无填充)返 null 跳过,不膨胀
    const style = extractStyle(cell, theme)
    const hasBorder = !!(style.borders.top || style.borders.bottom || style.borders.left || style.borders.right || style.borders.diagonal)
    const hasFill = style.fill.type !== 'none'
    if (!hasBorder && !hasFill) return null
    return { row, col, type: 'empty', raw: null, styleId: internStyle(style) }
  }

  let type: CellModel['type'] = 'string'
  let raw: CellModel['raw'] = null
  let rich: RichTextRun[] | undefined
  let formula: string | undefined
  let hyperlink: string | undefined

  if (typeof value === 'number') {
    type = 'number'
    raw = value
  } else if (typeof value === 'boolean') {
    type = 'boolean'
    raw = value
  } else if (value instanceof Date) {
    type = 'date'
    raw = value
  } else if (typeof value === 'string') {
    type = 'string'
    raw = value
  } else if (typeof value === 'object') {
    const v: any = value
    if (v.richText) {
      type = 'richtext'
      rich = (v.richText as any[]).map((r) => ({
        text: r.text ?? '',
        font: r.font ? toFont(r.font, theme) : undefined,
      }))
      raw = rich.map((r) => r.text).join('')
    } else if (v.formula !== undefined || v.sharedFormula !== undefined) {
      type = 'formula'
      formula = v.formula ?? v.sharedFormula
      const res = v.result
      if (res instanceof Date) raw = res
      else if (res && typeof res === 'object' && 'error' in res) raw = res.error
      else raw = res ?? null
    } else if (v.hyperlink !== undefined) {
      type = 'hyperlink'
      hyperlink = v.hyperlink
      raw = v.text ?? v.hyperlink
    } else if (v.error !== undefined) {
      type = 'error'
      raw = v.error
    } else {
      raw = String(v)
    }
  }

  const style = extractStyle(cell, theme)
  const comment = extractComment((cell as any).note)
  return { row, col, type, raw, rich, formula, hyperlink, comment, styleId: internStyle(style) }
}

/** ExcelJS note 可能是字符串或 { texts: [{text}] } 富文本，统一成纯文本 */
function extractComment(note: any): string | undefined {
  if (!note) return undefined
  if (typeof note === 'string') return note
  if (Array.isArray(note.texts)) {
    const s = note.texts.map((t: any) => t?.text ?? '').join('')
    return s || undefined
  }
  if (typeof note.text === 'string') return note.text
  return undefined
}

/**
 * OOXML 内置日期/时间格式(numFmtId 14-22 等)的"显示"本应跟随区域设置,但 ExcelJS 把它们**硬编码成美式串**
 * (如内置 14 → `mm-dd-yy` → 渲染 `04-01-26`)。WPS/Excel 中文环境显示的是 `2026/4/1`。这里把 ExcelJS 这些
 * 内置串重映射成中文 locale 习惯,跟 WPS 渲染对齐(本组件主要服务中文/WPS 文件)。
 * 仅命中 ExcelJS 内置串的"日期类",纯时间(h:mm 等)和带英文月名的(d-mmm-yy)保留不动。
 */
const EXCELJS_BUILTIN_DATE_LOCALE: Record<string, string> = {
  'mm-dd-yy': 'yyyy/m/d', // 内置 14 短日期
  'm/d/yy h:mm': 'yyyy/m/d h:mm', // 内置 22 日期+时间
}

function extractStyle(cell: ExcelJS.Cell, theme: CssColor[]): CellStyle {
  const s: any = cell.style || {}
  const align = s.alignment || {}
  const numFmt = s.numFmt || 'General'
  return {
    font: toFont(s.font, theme),
    fill: toFill(s.fill, theme),
    borders: toBorders(s.border, theme),
    hAlign: (align.horizontal as any) || 'general',
    vAlign: mapVAlign(align.vertical),
    wrapText: !!align.wrapText,
    shrinkToFit: !!align.shrinkToFit,
    // ExcelJS: 竖排返回 'vertical' → 255(渲染器据此走竖排堆叠);其余是 -90..90
    textRotation: align.textRotation === 'vertical' ? 255 : typeof align.textRotation === 'number' ? align.textRotation : 0,
    indent: align.indent || 0,
    numFmt: EXCELJS_BUILTIN_DATE_LOCALE[numFmt] ?? numFmt, // 内置短日期 locale 重映射(见上)
  }
}

function toFont(f: any, theme: CssColor[]): Font {
  if (!f) return { ...DEFAULT_FONT }
  return {
    name: f.name || DEFAULT_FONT.name,
    size: f.size || DEFAULT_FONT.size,
    bold: !!f.bold,
    italic: !!f.italic,
    underline: !!f.underline,
    strike: !!f.strike,
    color: resolveColor(f.color, theme) || '#000000',
  }
}

function toFill(fill: any, theme: CssColor[]): Fill {
  if (!fill || fill.type === 'none') return { type: 'none' }
  if (fill.type === 'pattern') {
    const pattern = fill.pattern
    if (pattern === 'none') return { type: 'none' }
    if (pattern === 'solid') {
      return {
        type: 'solid',
        fgColor: resolveColor(fill.fgColor, theme) || resolveColor(fill.bgColor, theme),
      }
    }
    return {
      type: 'pattern',
      pattern,
      fgColor: resolveColor(fill.fgColor, theme),
      bgColor: resolveColor(fill.bgColor, theme),
    }
  }
  if (fill.type === 'gradient') {
    const stops = (fill.stops || []).map((st: any) => ({
      position: st.position ?? 0,
      color: resolveColor(st.color, theme) || '#FFFFFF',
    }))
    // ExcelJS: gradient === 'angle' 时有 degree(度);'path' 型近似按对角
    const gradientAngle = typeof fill.degree === 'number' ? fill.degree : undefined
    return { type: 'gradient', gradientStops: stops, gradientAngle, fgColor: stops[0]?.color }
  }
  return { type: 'none' }
}

function toBorders(b: any, theme: CssColor[]) {
  if (!b) return {}
  const edge = (e: any) =>
    e && e.style
      ? { style: e.style, color: resolveColor(e.color, theme) || '#000000' }
      : undefined
  // ExcelJS: border.diagonal = { up, down, style, color }
  const diag = b.diagonal
  return {
    top: edge(b.top),
    bottom: edge(b.bottom),
    left: edge(b.left),
    right: edge(b.right),
    diagonal: edge(diag),
    diagonalUp: !!diag?.up,
    diagonalDown: !!diag?.down,
  }
}

function mapVAlign(v: any): CellStyle['vAlign'] {
  if (v === 'top') return 'top'
  if (v === 'middle') return 'middle'
  return 'bottom' // Excel 默认垂直靠下
}

function parseMerges(ws: ExcelJS.Worksheet): MergeRange[] {
  const model: any = (ws as any).model
  const raw: string[] = model?.merges || []
  const out: MergeRange[] = []
  for (const range of raw) {
    const m = parseA1Range(range)
    if (m) out.push(m)
  }
  return out
}

function parseConditional(ws: ExcelJS.Worksheet, theme: CssColor[]): ConditionalRule[] {
  const cfs: any[] = (ws as any).conditionalFormattings || []
  const out: ConditionalRule[] = []
  for (const cf of cfs) {
    const ranges = parseRefRanges(cf.ref)
    for (const rule of cf.rules || []) {
      const base: ConditionalRule = {
        ranges,
        priority: rule.priority ?? 0,
        type: 'unsupported',
      }
      switch (rule.type) {
        case 'cellIs':
          base.type = 'cellIs'
          base.operator = rule.operator
          base.formulae = rule.formulae
          base.style = rule.style ? cfStyle(rule.style, theme) : undefined
          break
        case 'expression':
          base.type = 'expression'
          base.formulae = rule.formulae
          base.style = rule.style ? cfStyle(rule.style, theme) : undefined
          break
        case 'colorScale': {
          base.type = 'colorScale'
          const colors = (rule.color || []).map((c: any) => resolveColor(c, theme) || '#FFFFFF')
          if (colors.length === 3) base.colorScale = { min: colors[0], mid: colors[1], max: colors[2] }
          else if (colors.length === 2) base.colorScale = { min: colors[0], max: colors[1] }
          break
        }
        case 'dataBar':
          base.type = 'dataBar'
          base.dataBar = {
            color: resolveColor(rule.color, theme) || '#638EC6',
            gradient: rule.gradient !== false,
          }
          break
        case 'iconSet':
          base.type = 'iconSet'
          base.iconSet = { name: rule.iconSet || '3TrafficLights1' }
          break
        case 'top10':
          base.type = 'top10'
          base.style = rule.style ? cfStyle(rule.style, theme) : undefined
          break
      }
      out.push(base)
    }
  }
  return out.sort((a, b) => a.priority - b.priority)
}

function cfStyle(s: any, theme: CssColor[]): Partial<CellStyle> {
  const out: Partial<CellStyle> = {}
  if (s.font) out.font = toFont(s.font, theme)
  if (s.fill) out.fill = toFill(s.fill, theme)
  if (s.border) out.borders = toBorders(s.border, theme)
  return out
}

function parseAutoFilter(ws: ExcelJS.Worksheet): MergeRange | undefined {
  const af: any = (ws as any).autoFilter
  if (!af) return undefined
  if (typeof af === 'string') return parseA1Range(af) || undefined
  if (af.from && af.to) {
    const from = typeof af.from === 'string' ? cellRef(af.from) : { row: af.from.row - 1, col: af.from.column - 1 }
    const to = typeof af.to === 'string' ? cellRef(af.to) : { row: af.to.row - 1, col: af.to.column - 1 }
    if (from && to) return { top: from.row, left: from.col, bottom: to.row, right: to.col }
  }
  return undefined
}

// ---- A1 解析 ----
export function colLettersToIndex(letters: string): number {
  let n = 0
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64)
  }
  return n - 1 // 0-based
}

function cellRef(a1: string): { row: number; col: number } | null {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a1.trim())
  if (!m) return null
  return { col: colLettersToIndex(m[1]), row: parseInt(m[2], 10) - 1 }
}

export function parseA1Range(range: string): MergeRange | null {
  const parts = range.split(':')
  const a = cellRef(parts[0])
  const b = parts[1] ? cellRef(parts[1]) : a
  if (!a || !b) return null
  return {
    top: Math.min(a.row, b.row),
    left: Math.min(a.col, b.col),
    bottom: Math.max(a.row, b.row),
    right: Math.max(a.col, b.col),
  }
}

/** ref 可能是 "A1:B2 C3:D4" 多段(空格分隔) */
function parseRefRanges(ref: string | undefined): MergeRange[] {
  if (!ref) return []
  return ref
    .split(/\s+/)
    .map((r) => parseA1Range(r))
    .filter((x): x is MergeRange => !!x)
}
