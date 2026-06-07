import { useCallback, useRef, useState } from 'react'
import type { TransformModelFn, WorkbookModel } from '@/core/model/types'
import type { ParseProgress } from '@/core/progress'
import { loadArrayBuffer, type ExcelSource } from '@/core/loader'
import { detectFormat, finalizeImages, friendlyError, revokeImages } from '@/core/finalize'
import { parseInWorker } from '@/composables/worker-client'

export interface ExcelDocument {
  loading: boolean
  error: string | null
  workbook: WorkbookModel | null
  progress: ParseProgress | null
  load: (src: ExcelSource, transform?: TransformModelFn) => Promise<void>
  /** 直接喂模型(JSON 直渲 / 模板已应用),跳过 parser */
  loadModel: (model: WorkbookModel, transform?: TransformModelFn) => void
  /** 原始字节(高保真 overlay 导出用) */
  sourceBuffer: ArrayBuffer | null
}

/**
 * React 版文档加载 hook —— 与 Vue 的 useExcelDocument 同逻辑(共用 core 的 loader/finalize/worker)。
 * 解析走 worker(大文件不卡),按文件头给友好错误,完成后 finalizeImages 落 blob url。
 */
export function useExcelDocument(): ExcelDocument {
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [workbook, setWorkbook] = useState<WorkbookModel | null>(null)
  const [progress, setProgress] = useState<ParseProgress | null>(null)
  const [sourceBuffer, setSourceBuffer] = useState<ArrayBuffer | null>(null)
  const wbRef = useRef<WorkbookModel | null>(null)

  const load = useCallback(async (src: ExcelSource, transform?: TransformModelFn) => {
    setLoading(true)
    setError(null)
    if (wbRef.current) revokeImages(wbRef.current) // 释放上一份图片 blob
    wbRef.current = null
    setWorkbook(null)
    setSourceBuffer(null)
    setProgress({ stage: 'read', ratio: 0 })
    try {
      const buffer = await loadArrayBuffer(src, (loaded, total) =>
        setProgress({ stage: 'read', ratio: total ? loaded / total : undefined }),
      )
      const fmt = detectFormat(buffer)
      if (fmt === 'xls') throw new Error('这是旧版 .xls(BIFF) 或加密文件，本预览器只支持 .xlsx/.xlsm。')
      if (fmt === 'not-zip') throw new Error('文件不是有效的 .xlsx(非 ZIP 包)。')
      if (fmt === 'empty') throw new Error('文件为空。')

      let model = await parseInWorker(buffer, (p) => setProgress(p))
      if (transform) model = transform(model) ?? model
      finalizeImages(model)
      setSourceBuffer(buffer.slice(0)) // 留一份原件副本(overlay 高保真导出用)
      wbRef.current = model
      setWorkbook(model)
    } catch (e: unknown) {
      const msg = (e as Error)?.message
      setError(msg && /预览器|损坏|加密|为空|ZIP/.test(msg) ? msg : friendlyError(e))
      console.error('[ooxml-preview] 解析失败:', e)
    } finally {
      setLoading(false)
      setProgress(null)
    }
  }, [])

  const loadModel = useCallback((model: WorkbookModel, transform?: TransformModelFn) => {
    if (wbRef.current) revokeImages(wbRef.current)
    setSourceBuffer(null)
    setError(null)
    setProgress(null)
    const m = transform ? (transform(model) ?? model) : model
    wbRef.current = m
    setWorkbook(m)
  }, [])

  return { loading, error, workbook, progress, load, loadModel, sourceBuffer }
}
