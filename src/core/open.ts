/**
 * `openWorkbook` —— 一行打开 .xlsx 的便捷门面(框架无关,浏览器 + 纯 Node 通用)。
 *
 * = `loadArrayBuffer(src)` 归一化输入 + `parseWorkbook(...)` 解析成中间模型。
 *
 * - **Node**:`await openWorkbook(fs.readFileSync('x.xlsx'))` —— Buffer 是 Uint8Array 子类,直接受理。
 * - **浏览器**:也接受 `File` / `Blob` / `ArrayBuffer` / `Uint8Array` / URL 字符串。
 *
 * 注:URL 字符串走 `fetch`,**Node 下不支持本地路径 / `file://`** —— Node 端请用 `fs` 读出 Buffer 再传。
 */
import type { WorkbookModel } from './model/types'
import type { ProgressFn } from './progress'
import { loadArrayBuffer, type ExcelSource } from './loader'
import { parseWorkbook } from './parser'

export async function openWorkbook(src: ExcelSource, onProgress?: ProgressFn): Promise<WorkbookModel> {
  const ab = await loadArrayBuffer(src)
  return parseWorkbook(ab, onProgress)
}
