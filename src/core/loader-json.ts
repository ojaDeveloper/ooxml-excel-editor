/**
 * JSON 直渲(P3)—— 不走 parser(无 .xlsx 文件),直接把 JSON 构造成 WorkbookModel,壳照常 render。
 *
 * 支持三种 input shape(自动识别):
 *  1. `unknown[][]`                 二维数组 → 单表,首格 A1
 *  2. `Record<string,unknown>[]`    对象数组 → 单表,首行表头(默认行为)= keys
 *  3. `{ sheets: [...] }`           多表 + 可指定名 + 各表 rows 仍走上述识别
 *
 * 类型自动推断(默认 on):数字字符串 → number,ISO 日期串 → Date,'TRUE'/'FALSE' → boolean。
 * 默认样式 = 单一缺省 CellStyle(makeDefaultStyle());壳的 :theme/:cellStyle/插件 仍可叠加。
 */
import type { CellModel, CellStyle, CellValueType, SheetModel, WorkbookModel } from './model/types'
import { cellKey } from './model/types'

export type JsonRow = unknown[] | Record<string, unknown>
export type JsonSheetInput = { name?: string; rows: JsonRow[] }
export type JsonInput =
  | unknown[][]
  | Record<string, unknown>[]
  | { sheets: JsonSheetInput[] }

export interface JsonLoadOptions {
  /** 对象数组:首行写表头(键名);默认 true。`false` 直接从首行起当数据(用 columns 控制顺序) */
  headerRow?: boolean
  /** 单表 / 默认表名,缺省 'Sheet1' */
  sheetName?: string
  /** 主题色 17 色调色板;不给用全黑兜底 */
  themeColors?: string[]
  /** 数字串 → 数字、'TRUE'/'FALSE' → boolean、ISO 日期串 → Date(默认 true) */
  autoInfer?: boolean
}

const DEFAULT_THEME: string[] = Array(17).fill('#000000')

/** 一份默认 CellStyle,与 ExcelJS 解析得到的"无样式"格视觉接近。 */
function makeDefaultStyle(): CellStyle {
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

function makeSheet(name: string, index: number): SheetModel {
  return {
    name,
    index,
    state: 'visible',
    dimension: { rows: 0, cols: 0 },
    cells: new Map(),
    styles: [makeDefaultStyle()],
    merges: [],
    columns: new Map(),
    rows: new Map(),
    defaultColWidth: 64,
    defaultRowHeight: 20,
    freeze: { frozenRows: 0, frozenCols: 0 },
    conditional: [],
    dataValidations: [],
    images: [],
    charts: [],
    shapes: [],
    sparklines: [],
    showGridLines: true,
  }
}

/** 推断单值:数字字符串 → number;'TRUE'/'FALSE' → boolean;ISO 日期串 → Date;其余原样。 */
function inferValue(v: unknown, autoInfer: boolean): { raw: CellModel['raw']; type: CellValueType } {
  if (v == null || v === '') return { raw: null, type: 'empty' }
  if (typeof v === 'number') return { raw: v, type: 'number' }
  if (typeof v === 'boolean') return { raw: v, type: 'boolean' }
  if (v instanceof Date) return { raw: v, type: 'date' }
  if (typeof v !== 'string') return { raw: String(v), type: 'string' }

  if (!autoInfer) return { raw: v, type: 'string' }

  // 数字推断:整数 / 小数 / 科学计数 / 带千分位(去掉再 parse)
  const trimmed = v.trim()
  if (/^[+-]?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(trimmed)) {
    const n = Number(trimmed)
    if (Number.isFinite(n)) return { raw: n, type: 'number' }
  }
  // boolean
  if (/^(TRUE|FALSE)$/i.test(trimmed)) return { raw: /^TRUE$/i.test(trimmed), type: 'boolean' }
  // ISO 日期(YYYY-MM-DD[ T...])
  if (/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?$/.test(trimmed)) {
    const d = new Date(trimmed)
    if (!isNaN(d.getTime())) return { raw: d, type: 'date' }
  }
  return { raw: v, type: 'string' }
}

/** 一行对象 → 按 keys 顺序的值数组(对象数组用) */
function objectRowToArray(obj: Record<string, unknown>, keys: string[]): unknown[] {
  return keys.map((k) => obj[k])
}

/** 把 rows(2D 或对象数组)写入 sheet,从 startRow 起。返回写入后的最大 row + col。 */
function writeRows(sheet: SheetModel, rows: JsonRow[], opts: { startRow?: number; startCol?: number; headerRow?: boolean; columns?: string[]; autoInfer?: boolean } = {}): { lastRow: number; lastCol: number; wrote: boolean } {
  const startRow = opts.startRow ?? 0
  const startCol = opts.startCol ?? 0
  const autoInfer = opts.autoInfer ?? true
  let maxRow = 0
  let maxCol = 0
  let wrote = false
  if (!rows.length) return { lastRow: 0, lastCol: 0, wrote: false }
  const first = rows[0]
  const isObjectArray = first !== null && typeof first === 'object' && !Array.isArray(first)
  // 对象数组列序:opts.columns > 首行 keys
  const keys: string[] = isObjectArray ? (opts.columns ?? Object.keys(first as Record<string, unknown>)) : []

  let writeRowIdx = startRow
  // 对象数组且 headerRow=true 时,先写一行表头
  if (isObjectArray && (opts.headerRow ?? true)) {
    for (let c = 0; c < keys.length; c++) {
      const r = writeRowIdx
      const col = startCol + c
      sheet.cells.set(cellKey(r, col), { row: r, col, type: 'string', raw: keys[c], styleId: 0 })
      maxRow = Math.max(maxRow, r)
      maxCol = Math.max(maxCol, col)
      wrote = true
    }
    writeRowIdx++
  }

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]
    const arr: unknown[] = Array.isArray(row) ? row : objectRowToArray(row as Record<string, unknown>, keys)
    for (let c = 0; c < arr.length; c++) {
      const r = writeRowIdx + i - (isObjectArray && (opts.headerRow ?? true) ? 0 : 0)
      const col = startCol + c
      const inf = inferValue(arr[c], autoInfer)
      if (inf.type === 'empty') continue
      sheet.cells.set(cellKey(r, col), { row: r, col, type: inf.type, raw: inf.raw, styleId: 0 })
      maxRow = Math.max(maxRow, r)
      maxCol = Math.max(maxCol, col)
      wrote = true
    }
  }
  return { lastRow: maxRow, lastCol: maxCol, wrote }
}

/** JSON → WorkbookModel。input shape 自动识别。 */
export function jsonToWorkbook(input: JsonInput, opts: JsonLoadOptions = {}): WorkbookModel {
  const themeColors = (opts.themeColors && opts.themeColors.length === 17) ? opts.themeColors : DEFAULT_THEME
  const autoInfer = opts.autoInfer ?? true
  const headerRow = opts.headerRow ?? true

  // 多表 shape:{ sheets: [...] }
  if (input && typeof input === 'object' && !Array.isArray(input) && Array.isArray((input as { sheets: unknown }).sheets)) {
    const multi = (input as { sheets: JsonSheetInput[] }).sheets
    const sheets: SheetModel[] = multi.map((s, i) => {
      const sheet = makeSheet(s.name || `Sheet${i + 1}`, i)
      const { lastRow, lastCol, wrote } = writeRows(sheet, s.rows ?? [], { headerRow, autoInfer })
      sheet.dimension = wrote ? { rows: lastRow + 1, cols: lastCol + 1 } : { rows: 0, cols: 0 }
      return sheet
    })
    return { sheets, activeSheet: 0, themeColors, date1904: false }
  }

  // 单表:二维数组 / 对象数组
  const sheet = makeSheet(opts.sheetName || 'Sheet1', 0)
  const rows = (input as JsonRow[]) ?? []
  const { lastRow, lastCol, wrote } = writeRows(sheet, rows, { headerRow, autoInfer })
  sheet.dimension = wrote ? { rows: lastRow + 1, cols: lastCol + 1 } : { rows: 0, cols: 0 }
  return { sheets: [sheet], activeSheet: 0, themeColors, date1904: false }
}

/** 用方传 WorkbookModel-shape 时的浅校验(壳分支用) */
export function isWorkbookModel(v: unknown): v is WorkbookModel {
  return !!v && typeof v === 'object' && Array.isArray((v as { sheets: unknown }).sheets) &&
    typeof (v as WorkbookModel).activeSheet === 'number' &&
    Array.isArray((v as WorkbookModel).themeColors)
}
