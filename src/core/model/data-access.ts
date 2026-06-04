/**
 * 数据访问 API —— 在中间模型(WorkbookModel/SheetModel)之上提供"好用"的读取接口。
 * 独立函数,配 parseWorkbook 即可用(不必渲染);组件 ref 方法是它们的薄封装。
 *
 * 值 vs 文本: format 默认 true → 套了数字/日期格式的"显示文本"(所见即所得);
 *            format:false → 原始值(number/boolean/Date/string/null)。
 * 合并单元格: 覆盖格在模型里是空,2D/JSON 里由锚点(左上)持值、其余为空 —— 同 Excel 直觉。
 * 公式: 不重算,沿用 Excel 缓存结果(值在 raw / 文本里;公式串见 cell.formula)。
 */
import type { CellModel, CellStyle, MergeRange, SheetModel, WorkbookModel } from './types'
import { cellKey } from './types'
import { formatValue } from '../format/number-format'
import { colIndexToLetters } from '../layout/grid-metrics'

export type CellValue = number | string | boolean | Date | null

export interface ReadOptions {
  /** true(默认)=格式化显示文本;false=原始值 */
  format?: boolean
  /** 1904 日期系统(组件 ref 方法会自动带上;独立调用时从 workbook.date1904 取) */
  date1904?: boolean
}

/** 单元格显示文本(套数字/日期格式;富文本拼接)。纯函数,渲染器与数据 API 共用。 */
export function cellDisplayText(
  cell: CellModel | undefined,
  style: CellStyle | undefined,
  date1904: boolean,
): string {
  if (!cell || cell.type === 'empty') return ''
  if (cell.type === 'richtext' && cell.rich) return cell.rich.map((r) => r.text).join('')
  return formatValue(cell.raw, style?.numFmt, date1904).text
}

/** 取单元格模型(Map 之上的便捷封装) */
export function getCell(sheet: SheetModel, row: number, col: number): CellModel | undefined {
  return sheet.cells.get(cellKey(row, col))
}

/** 取单元格原始值(number/boolean/Date/string/null) */
export function getCellValue(sheet: SheetModel, row: number, col: number): CellValue {
  const cell = getCell(sheet, row, col)
  if (!cell || cell.type === 'empty') return null
  return cell.raw
}

/** 取单元格解析后的样式(经 styleId);无则 undefined */
export function getCellStyle(sheet: SheetModel, row: number, col: number): CellStyle | undefined {
  const cell = getCell(sheet, row, col)
  if (!cell) return undefined
  return sheet.styles[cell.styleId]
}

/** 取单元格格式化显示文本 */
export function getCellText(sheet: SheetModel, row: number, col: number, date1904 = false): string {
  const cell = getCell(sheet, row, col)
  return cellDisplayText(cell, cell ? sheet.styles[cell.styleId] : undefined, date1904)
}

/** 合并区"非锚点覆盖格"集合(解析器会把合并值灌进每个覆盖格,这里据此清空,只让锚点持值) */
function coveredCells(merges: MergeRange[]): Set<string> {
  const set = new Set<string>()
  for (const m of merges) {
    for (let r = m.top; r <= m.bottom; r++) {
      for (let c = m.left; c <= m.right; c++) {
        if (!(r === m.top && c === m.left)) set.add(cellKey(r, c))
      }
    }
  }
  return set
}

/** 单格的值或文本(按 opts.format);covered=合并覆盖格集合,命中则返回空 */
function cellOut(
  sheet: SheetModel,
  row: number,
  col: number,
  format: boolean,
  date1904: boolean,
  covered?: Set<string>,
): CellValue {
  if (covered && covered.has(cellKey(row, col))) return format ? '' : null
  if (format) return getCellText(sheet, row, col, date1904)
  return getCellValue(sheet, row, col)
}

/** 整表 → 稠密二维数组(dimension.rows × dimension.cols)。空格为 ''(文本)/ null(值)。 */
export function getSheetData(sheet: SheetModel, opts: ReadOptions = {}): CellValue[][] {
  const format = opts.format ?? true
  const date1904 = opts.date1904 ?? false
  const covered = coveredCells(sheet.merges)
  const out: CellValue[][] = []
  for (let r = 0; r < sheet.dimension.rows; r++) {
    const line: CellValue[] = []
    for (let c = 0; c < sheet.dimension.cols; c++) line.push(cellOut(sheet, r, c, format, date1904, covered))
    out.push(line)
  }
  return out
}

/** 区域 → 二维数组(行×列,0-based 闭区间) */
export function getRangeData(sheet: SheetModel, range: MergeRange, opts: ReadOptions = {}): CellValue[][] {
  const format = opts.format ?? true
  const date1904 = opts.date1904 ?? false
  const covered = coveredCells(sheet.merges)
  const out: CellValue[][] = []
  for (let r = range.top; r <= range.bottom; r++) {
    const line: CellValue[] = []
    for (let c = range.left; c <= range.right; c++) line.push(cellOut(sheet, r, c, format, date1904, covered))
    out.push(line)
  }
  return out
}

export interface SheetToJSONOptions extends ReadOptions {
  /** 表头所在行(0-based,默认 0);数据从其下一行起 */
  headerRow?: number
}

/** 整表 → 对象数组(首行作 key)。空表头回退列字母;整行皆空则跳过。 */
export function sheetToJSON(sheet: SheetModel, opts: SheetToJSONOptions = {}): Record<string, CellValue>[] {
  const format = opts.format ?? true
  const date1904 = opts.date1904 ?? false
  const headerRow = opts.headerRow ?? 0
  const cols = sheet.dimension.cols
  const covered = coveredCells(sheet.merges)
  const headers: string[] = []
  for (let c = 0; c < cols; c++) {
    const h = cellOut(sheet, headerRow, c, true, date1904, covered) as string
    headers.push(h || colIndexToLetters(c))
  }
  const rows: Record<string, CellValue>[] = []
  for (let r = headerRow + 1; r < sheet.dimension.rows; r++) {
    let allEmpty = true
    const obj: Record<string, CellValue> = {}
    for (let c = 0; c < cols; c++) {
      const v = cellOut(sheet, r, c, format, date1904, covered)
      if (v !== null && v !== '') allEmpty = false
      obj[headers[c]] = v
    }
    if (!allEmpty) rows.push(obj)
  }
  return rows
}

/** 全簿 → { 表名: 对象数组 }(可见表;自动用 workbook.date1904) */
export function getWorkbookJSON(
  workbook: WorkbookModel,
  opts: SheetToJSONOptions = {},
): Record<string, Record<string, CellValue>[]> {
  const out: Record<string, Record<string, CellValue>[]> = {}
  for (const sheet of workbook.sheets) {
    if (sheet.state !== 'visible') continue
    out[sheet.name] = sheetToJSON(sheet, { ...opts, date1904: workbook.date1904 })
  }
  return out
}
