/** 解析入口: ArrayBuffer → WorkbookModel。 */
import type { WorkbookModel } from '../model/types'
import type { ProgressFn } from '../progress'
import { openPackage } from './raw-xml'
import { parseTheme } from './theme'
import { buildSheets } from './exceljs-adapter'
import { attachDrawings } from './drawing-parser'
import { attachCellImages } from './cell-image-parser'
import { attachRowMeta } from './row-meta-parser'
import { attachSparklines } from './sparkline-parser'
import { attachPageBreaks } from './page-break-parser'
import { attachPivotTables } from './pivot-parser'

export async function parseWorkbook(
  buffer: ArrayBuffer | Uint8Array,
  onProgress?: ProgressFn,
): Promise<WorkbookModel> {
  // 归一化输入: 接受 ArrayBuffer 或 Uint8Array(含 Node Buffer —— 它是 Uint8Array 子类),
  // 便于纯 Node 直接传 fs.readFileSync() 的结果。ArrayBuffer 老调用 100% 兼容。
  const buf: ArrayBuffer =
    buffer instanceof Uint8Array
      ? (buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer)
      : buffer

  // 1. 原始包(用于 theme / drawings / charts) —— 复用一份 buffer
  const pkg = openPackage(buf.slice(0))
  const themeColors = parseTheme(pkg)

  // 2. ExcelJS 主解析(黑盒,无法报进度 → 进入不确定态)
  onProgress?.({ stage: 'parse' })
  const ExcelJS = (await import('exceljs')).default
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buf)

  // 3. 构建中间模型(我们自己的遍历,可报真实进度)
  const sheets = buildSheets(wb, themeColors, onProgress)

  // 3. 图片 / 图表锚点补齐
  try {
    attachDrawings(pkg, sheets)
  } catch (e) {
    console.warn('[ooxml-preview] drawings 解析失败，跳过图片/图表:', e)
  }

  // 3.5 WPS 单元格内嵌图(DISPIMG):登记表 + 单元格 dispImgId 标记
  let cellImages: WorkbookModel['cellImages']
  try {
    cellImages = attachCellImages(pkg, sheets)
  } catch (e) {
    console.warn('[ooxml-preview] WPS 单元格内嵌图(DISPIMG)解析失败，跳过:', e)
  }

  // 4. 迷你图(sparklines)补齐
  try {
    attachSparklines(pkg, sheets)
  } catch (e) {
    console.warn('[ooxml-preview] sparklines 解析失败，跳过迷你图:', e)
  }

  // 4.5 透视表只读 UI 元数据(字段按钮/范围);数据仍由 worksheet 普通单元格显示
  try {
    attachPivotTables(pkg, sheets)
  } catch (e) {
    console.warn('[ooxml-preview] 透视表 UI 元数据解析失败，跳过:', e)
  }

  // 5. 手动分页符
  try {
    attachPageBreaks(pkg, sheets)
  } catch (e) {
    console.warn('[ooxml-preview] 分页符解析失败，跳过:', e)
  }

  // 6. 行 customHeight 标记(决定渲染层是否对该行做自动行高)
  try {
    attachRowMeta(pkg, sheets)
  } catch (e) {
    console.warn('[ooxml-preview] 行高标记解析失败，跳过:', e)
  }

  const date1904 = !!(wb as any).properties?.date1904

  return {
    sheets,
    activeSheet: 0,
    themeColors,
    date1904,
    cellImages,
  }
}
