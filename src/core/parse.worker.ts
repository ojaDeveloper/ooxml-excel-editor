/**
 * 解析 Worker: 把 exceljs + 原始 XML 解析放到后台线程,大文件不卡 UI。
 * 入: ArrayBuffer。出: 多条 {type:'progress'} + 一条 {type:'done'}。
 * 图片以字节随模型传回(结构化克隆,不 transfer 以免重复 ArrayBuffer 报错)。
 */
import { parseWorkbook } from './parser'
import type { WorkbookModel } from './model/types'
import type { ParseProgress } from './progress'

export type WorkerMsg =
  | { type: 'progress'; progress: ParseProgress }
  | { type: 'done'; ok: true; model: WorkbookModel }
  | { type: 'done'; ok: false; error: string }

const post = (msg: WorkerMsg) => (self as unknown as Worker).postMessage(msg)

self.onmessage = async (e: MessageEvent<ArrayBuffer>) => {
  try {
    const model = await parseWorkbook(e.data, (progress) => post({ type: 'progress', progress }))
    post({ type: 'done', ok: true, model })
  } catch (err) {
    post({ type: 'done', ok: false, error: (err as any)?.message || String(err) })
  }
}
