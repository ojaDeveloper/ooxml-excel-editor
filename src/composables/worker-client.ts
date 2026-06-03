/**
 * Web Worker 解析客户端(dev / demo 用)。
 * 库构建时本模块被 vite 别名替换为 worker-client.stub.ts(纯主线程),
 * 以免把 exceljs 随预打包 worker 打进库。
 */
import type { WorkbookModel } from '@/core/model/types'
import type { WorkerMsg } from '@/core/parse.worker'
import type { ProgressFn } from '@/core/progress'
import { parseWorkbook } from '@/core/parser'

export function parseInWorker(buffer: ArrayBuffer, onProgress?: ProgressFn): Promise<WorkbookModel> {
  let worker: Worker
  try {
    worker = new Worker(new URL('../core/parse.worker.ts', import.meta.url), { type: 'module' })
  } catch {
    return parseWorkbook(buffer, onProgress) // 环境不支持 module worker → 主线程
  }
  return new Promise<WorkbookModel>((resolve, reject) => {
    worker.onmessage = (ev: MessageEvent<WorkerMsg>) => {
      const m = ev.data
      if (m.type === 'progress') {
        onProgress?.(m.progress)
        return
      }
      worker.terminate()
      if (m.ok) resolve(m.model)
      else reject(new Error(m.error)) // 文件解析错，照实抛
    }
    worker.onerror = () => {
      worker.terminate()
      console.warn('[ooxml-preview] worker 不可用，回退主线程解析')
      resolve(parseWorkbook(buffer, onProgress)) // 加载/运行失败 → 主线程兜底
    }
    worker.postMessage(buffer) // 不 transfer，保留原件以便回退
  })
}
