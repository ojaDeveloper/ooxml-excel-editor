<script setup lang="ts">
/**
 * 内置导出进度遮罩(P1.5)—— 居中模态,暗背景 + 白卡 + stage 标签 + 进度条 + 取消。
 * 默认在调 viewer.downloadPdf / exportImage / exportXlsx / print / 选区图片批量转换时自动显示。
 * 用方覆盖路径见 ExcelViewer.vue 的 `:export-progress` prop 与 `#export-progress` slot。
 */
import type { ExportProgress } from '@/core/progress'
defineProps<{ state: ExportProgress | null; busy: boolean }>()
const emit = defineEmits<{ (e: 'cancel'): void }>()
const stageLabel: Record<string, string> = {
  render: '渲染中',
  compose: '合成中',
  paginate: '分页中',
  write: '写出文件',
  zip: 'zip 压缩',
  convert: '批量转换',
}
</script>

<template>
  <Teleport to="body">
    <div v-if="busy" class="export-progress-overlay" role="dialog" aria-modal="true" aria-live="polite">
      <div class="card">
        <div class="title">{{ state?.label || stageLabel[state?.stage ?? ''] || '处理中…' }}</div>
        <div class="bar" :class="{ indeterminate: state?.ratio == null }">
          <div v-if="state?.ratio != null" class="fill" :style="{ width: Math.round((state.ratio ?? 0) * 100) + '%' }" />
        </div>
        <div class="row">
          <span class="pct">{{ state?.ratio != null ? Math.round(state.ratio * 100) + '%' : '正在处理…' }}</span>
          <button class="cancel" @click="emit('cancel')" title="按 Esc 也可取消">取消</button>
        </div>
      </div>
    </div>
  </Teleport>
</template>

<style scoped>
.export-progress-overlay {
  position: fixed; inset: 0; z-index: 9999;
  background: rgba(0, 0, 0, 0.42);
  display: flex; align-items: center; justify-content: center;
  font: 13px/1.5 -apple-system, 'Segoe UI', sans-serif;
}
.card {
  min-width: 320px; max-width: 480px;
  background: #fff; border-radius: 10px;
  box-shadow: 0 10px 32px rgba(0, 0, 0, 0.28);
  padding: 18px 20px;
}
.title { font-size: 14px; font-weight: 600; color: #1f2329; margin-bottom: 12px; }
.bar { position: relative; height: 6px; background: #eceef1; border-radius: 3px; overflow: hidden; margin-bottom: 10px; }
.bar .fill { height: 100%; background: #21a366; transition: width 120ms linear; }
.bar.indeterminate::after {
  content: ''; position: absolute; inset: 0;
  background: linear-gradient(90deg, transparent, #21a366 40%, #21a366 60%, transparent);
  animation: ind 1.2s infinite;
}
@keyframes ind { from { transform: translateX(-100%); } to { transform: translateX(100%); } }
.row { display: flex; align-items: center; justify-content: space-between; }
.pct { color: #6b7280; font-size: 12px; }
.cancel {
  border: 1px solid #d8dbde; background: #fff; color: #1f2329;
  padding: 5px 14px; border-radius: 5px; cursor: pointer; font-size: 13px;
}
.cancel:hover { background: #f5f6f7; }
</style>
