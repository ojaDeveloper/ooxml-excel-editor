/**
 * Vue 2 壳的文档加载 composable — 跟 Vue 3 / React 壳同逻辑, 共用 core 的 loader/finalize/worker.
 * 用 Vue 2.7 内置 Composition API (ref).
 */
// Composition API 从 @vue/composition-api 拿 → 兼容 Vue 2.6.x + 2.7+
//   dev 时 vite alias 重定向到 vue2 (vue@2.7 dist); build 时 rollup external
import { ref } from '@vue/composition-api'
import type { TransformModelFn, WorkbookModel } from '@/core/model/types'
import type { ParseProgress } from '@/core/progress'
import { loadArrayBuffer, type ExcelSource } from '@/core/loader'
import { detectFormat, finalizeImages, friendlyError, revokeImages } from '@/core/finalize'
import { parseInWorker } from '@/composables/worker-client'

export function useExcelDocumentVue2() {
  const loading = ref(false)
  const error = ref<string | null>(null)
  const workbook = ref<WorkbookModel | null>(null) as { value: WorkbookModel | null }
  const progress = ref<ParseProgress | null>(null) as { value: ParseProgress | null }
  const sourceBuffer = ref<ArrayBuffer | null>(null) as { value: ArrayBuffer | null }

  async function load(src: ExcelSource, transform?: TransformModelFn) {
    loading.value = true
    error.value = null
    if (workbook.value) revokeImages(workbook.value)
    workbook.value = null
    sourceBuffer.value = null
    progress.value = { stage: 'read', ratio: 0 }
    try {
      const buffer = await loadArrayBuffer(src, (loaded, total) => {
        progress.value = { stage: 'read', ratio: total ? loaded / total : undefined }
      })
      const fmt = detectFormat(buffer)
      if (fmt === 'xls') throw new Error('这是旧版 .xls(BIFF) 或加密文件,本预览器只支持 .xlsx/.xlsm.')
      if (fmt === 'not-zip') throw new Error('文件不是有效的 .xlsx(非 ZIP 包).')
      if (fmt === 'empty') throw new Error('文件为空.')

      let model = await parseInWorker(buffer, (p) => { progress.value = p })
      if (transform) model = transform(model) ?? model
      finalizeImages(model)
      sourceBuffer.value = buffer.slice(0)
      workbook.value = model
    } catch (e: unknown) {
      const msg = (e as Error)?.message
      error.value = msg && /预览器|损坏|加密|为空|ZIP/.test(msg) ? msg : friendlyError(e)
      console.error('[ooxml-preview vue2] 解析失败:', e)
    } finally {
      loading.value = false
      progress.value = null
    }
  }

  function loadModel(model: WorkbookModel, transform?: TransformModelFn) {
    if (workbook.value) revokeImages(workbook.value)
    sourceBuffer.value = null
    error.value = null
    progress.value = null
    workbook.value = transform ? (transform(model) ?? model) : model
  }

  return { loading, error, workbook, load, loadModel, progress, sourceBuffer }
}
