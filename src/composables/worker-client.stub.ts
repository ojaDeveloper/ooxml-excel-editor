/**
 * worker-client 的库构建替身: 不引入任何 Worker(也就不会把 exceljs 预打包进库),
 * 直接主线程解析。需要 Worker 的应用可用导出的 parseWorkbook 自行包 Worker。
 */
import type { WorkbookModel } from '@/core/model/types'
import type { ProgressFn } from '@/core/progress'
import { parseWorkbook } from '@/core/parser'

export function parseInWorker(buffer: ArrayBuffer, onProgress?: ProgressFn): Promise<WorkbookModel> {
  return parseWorkbook(buffer, onProgress)
}
