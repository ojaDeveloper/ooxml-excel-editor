/**
 * 数据导出(框架无关)—— JSON / CSV 直接建在读数据层(getSheetData/getWorkbookJSON)之上,
 * 零新值/格式逻辑 → 与渲染所见、与 XLSX 导出天然一致(要求 4:一份数据层喂所有格式)。
 */
import type { SheetModel, WorkbookModel } from '../model/types'
import { getSheetData, getWorkbookJSON, type CellValue, type SheetToJSONOptions } from '../model/data-access'

/** CSV 单元转义:含逗号/引号/换行则包引号,内部引号翻倍。 */
function csvCell(v: CellValue): string {
  if (v == null) return ''
  const s = v instanceof Date ? v.toISOString() : String(v)
  return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s
}

/** 一张表 → CSV 文本(默认格式化显示值,与渲染一致;行用 CRLF)。 */
export function toCsv(sheet: SheetModel, opts: { format?: boolean; date1904?: boolean } = {}): string {
  const data = getSheetData(sheet, { format: opts.format ?? true, date1904: opts.date1904 ?? false })
  return data.map((row) => row.map(csvCell).join(',')).join('\r\n')
}

/** 整簿 → JSON(各表首行作 key 的对象数组);复用 getWorkbookJSON。默认 raw 类型值(非显示串)。 */
export function toWorkbookJson(workbook: WorkbookModel, opts: SheetToJSONOptions = {}): string {
  return JSON.stringify(getWorkbookJSON(workbook, { format: false, date1904: workbook.date1904, ...opts }), null, 2)
}
