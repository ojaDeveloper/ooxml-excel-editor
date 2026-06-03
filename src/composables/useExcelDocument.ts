import { ref, shallowRef } from 'vue'
import type { WorkbookModel } from '@/core/model/types'
import type { ParseProgress } from '@/core/progress'
import { loadArrayBuffer, type ExcelSource } from '@/core/loader'
import { detectFormat, finalizeImages, friendlyError, revokeImages } from '@/core/finalize'
import { parseInWorker } from './worker-client'

export function useExcelDocument() {
  const loading = ref(false)
  const error = ref<string | null>(null)
  const workbook = shallowRef<WorkbookModel | null>(null)
  const progress = ref<ParseProgress | null>(null)

  async function load(src: ExcelSource) {
    loading.value = true
    error.value = null
    if (workbook.value) revokeImages(workbook.value) // 释放上一份图片 blob
    workbook.value = null
    progress.value = { stage: 'read', ratio: 0 }
    try {
      const buffer = await loadArrayBuffer(src, (loaded, total) => {
        progress.value = { stage: 'read', ratio: total ? loaded / total : undefined }
      })

      // 解析前先按文件头给出友好提示，避免对损坏/旧格式硬解析
      const fmt = detectFormat(buffer)
      if (fmt === 'xls') throw new Error('这是旧版 .xls(BIFF) 或加密文件，本预览器只支持 .xlsx/.xlsm。')
      if (fmt === 'not-zip') throw new Error('文件不是有效的 .xlsx(非 ZIP 包)。')
      if (fmt === 'empty') throw new Error('文件为空。')

      const model = await parseInWorker(buffer, (p) => {
        progress.value = p
      })
      finalizeImages(model)
      workbook.value = model
    } catch (e: any) {
      // detectFormat 抛的已是友好文案，直接用;底层异常再翻译
      error.value = e?.message && /预览器|损坏|加密|为空|ZIP/.test(e.message) ? e.message : friendlyError(e)
      console.error('[ooxml-preview] 解析失败:', e)
    } finally {
      loading.value = false
      progress.value = null
    }
  }

  return { loading, error, workbook, load, progress }
}
