/** 解析入口: ArrayBuffer → WorkbookModel。 */
import type { WorkbookModel } from '../model/types'
import type { ProgressFn } from '../progress'
import { openPackage } from './raw-xml'
import { parseTheme } from './theme'
import { buildSheets } from './exceljs-adapter'
import { attachDrawings } from './drawing-parser'
import { attachSparklines } from './sparkline-parser'
import { attachPageBreaks } from './page-break-parser'

export async function parseWorkbook(buffer: ArrayBuffer, onProgress?: ProgressFn): Promise<WorkbookModel> {
  // 1. 原始包(用于 theme / drawings / charts) —— 复用一份 buffer
  const pkg = openPackage(buffer.slice(0))
  const themeColors = parseTheme(pkg)

  // 2. ExcelJS 主解析(黑盒,无法报进度 → 进入不确定态)
  onProgress?.({ stage: 'parse' })
  const ExcelJS = (await import('exceljs')).default
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buffer)

  // 3. 构建中间模型(我们自己的遍历,可报真实进度)
  const sheets = buildSheets(wb, themeColors, onProgress)

  // 3. 图片 / 图表锚点补齐
  try {
    attachDrawings(pkg, sheets)
  } catch (e) {
    console.warn('[ooxml-preview] drawings 解析失败，跳过图片/图表:', e)
  }

  // 4. 迷你图(sparklines)补齐
  try {
    attachSparklines(pkg, sheets)
  } catch (e) {
    console.warn('[ooxml-preview] sparklines 解析失败，跳过迷你图:', e)
  }

  // 5. 手动分页符
  try {
    attachPageBreaks(pkg, sheets)
  } catch (e) {
    console.warn('[ooxml-preview] 分页符解析失败，跳过:', e)
  }

  const date1904 = !!(wb as any).properties?.date1904

  return {
    sheets,
    activeSheet: 0,
    themeColors,
    date1904,
  }
}
