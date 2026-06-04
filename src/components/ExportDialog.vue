<script setup lang="ts">
/** 导出设置对话框: 选范围(选区/当前表/全部表)+ 清晰度/表头/网格线 + 纸张方向(PDF/打印)。 */
import { computed, reactive } from 'vue'
import type { MergeRange } from '@/core/model/types'
import { colIndexToLetters } from '@/core/layout/grid-metrics'
import type { ExportConfig } from './export-types'

const props = defineProps<{
  selection: MergeRange | null
  sheetCount: number
}>()
const emit = defineEmits<{
  (e: 'close'): void
  (e: 'export', cfg: ExportConfig): void
}>()

const hasSelection = computed(() => {
  const s = props.selection
  return !!s && !(s.top === s.bottom && s.left === s.right)
})
const selectionLabel = computed(() => {
  const s = props.selection
  if (!s) return ''
  return `${colIndexToLetters(s.left)}${s.top + 1}:${colIndexToLetters(s.right)}${s.bottom + 1}`
})

const form = reactive<ExportConfig>({
  action: 'png',
  scope: hasSelection.value ? 'selection' : 'sheet',
  scale: 2,
  includeHeaders: false,
  gridlines: true,
  format: 'auto',
  orientation: 'auto',
  fitToWidth: true,
  pdfVector: false,
})

function run(action: ExportConfig['action']) {
  emit('export', { ...form, action })
}
</script>

<template>
  <div class="dlg-mask" @click.self="emit('close')">
    <div class="dlg" role="dialog" aria-label="导出设置">
      <div class="dlg-head">
        <span>导出 / 打印设置</span>
        <button class="x" @click="emit('close')" title="关闭">×</button>
      </div>

      <div class="dlg-body">
        <div class="field">
          <label class="lbl">范围</label>
          <div class="opts">
            <label :class="{ disabled: !hasSelection }">
              <input type="radio" value="selection" v-model="form.scope" :disabled="!hasSelection" />
              当前选区 <span v-if="hasSelection" class="hint">{{ selectionLabel }}</span>
              <span v-else class="hint">(未选多格)</span>
            </label>
            <label><input type="radio" value="sheet" v-model="form.scope" /> 当前工作表</label>
            <label><input type="radio" value="all" v-model="form.scope" /> 全部工作表 ({{ sheetCount }})</label>
          </div>
        </div>

        <div class="field">
          <label class="lbl">清晰度</label>
          <select v-model.number="form.scale">
            <option :value="1">标准 (1×)</option>
            <option :value="2">高清 (2×)</option>
            <option :value="3">超清 (3×)</option>
          </select>
        </div>

        <div class="field">
          <label class="lbl">内容</label>
          <div class="opts inline">
            <label><input type="checkbox" v-model="form.includeHeaders" /> 含行列号</label>
            <label><input type="checkbox" v-model="form.gridlines" /> 网格线</label>
          </div>
        </div>

        <div class="field">
          <label class="lbl">PDF 类型</label>
          <div class="opts">
            <label><input type="radio" :value="false" v-model="form.pdfVector" /> 位图 <span class="hint">(完整还原观感)</span></label>
            <label><input type="radio" :value="true" v-model="form.pdfVector" /> 矢量 <span class="hint">(文字可选可搜·清晰·文件小;中文需注册字体,否则该格转图)</span></label>
          </div>
        </div>

        <div class="field">
          <label class="lbl">纸张 <span class="hint">(PDF/打印)</span></label>
          <div class="opts inline">
            <select v-model="form.format">
              <option value="auto">自动(跟随表)</option>
              <option value="a4">A4</option>
              <option value="a3">A3</option>
              <option value="letter">Letter</option>
            </select>
            <select v-model="form.orientation">
              <option value="auto">方向: 自动</option>
              <option value="portrait">纵向</option>
              <option value="landscape">横向</option>
            </select>
            <label><input type="checkbox" v-model="form.fitToWidth" /> 适应页宽</label>
          </div>
        </div>
      </div>

      <div class="dlg-foot">
        <button class="ghost" @click="emit('close')">取消</button>
        <div class="grow" />
        <button @click="run('png')">导出 PNG</button>
        <button @click="run('pdf')">导出 PDF</button>
        <button class="primary" @click="run('print')">打印…</button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.dlg-mask {
  position: absolute;
  inset: 0;
  z-index: 30;
  background: rgba(0, 0, 0, 0.32);
  display: flex;
  align-items: center;
  justify-content: center;
}
.dlg {
  width: 420px;
  max-width: 92%;
  background: #fff;
  border-radius: 10px;
  box-shadow: 0 12px 40px rgba(0, 0, 0, 0.25);
  font-size: 13px;
  color: #333;
  overflow: hidden;
}
.dlg-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px 16px;
  font-weight: 600;
  border-bottom: 1px solid #eef0f2;
}
.dlg-head .x {
  border: none;
  background: none;
  font-size: 20px;
  line-height: 1;
  cursor: pointer;
  color: #999;
}
.dlg-head .x:hover { color: #444; }
.dlg-body { padding: 8px 16px 4px; }
.field {
  display: flex;
  align-items: flex-start;
  gap: 12px;
  padding: 9px 0;
  border-bottom: 1px solid #f4f5f6;
}
.field:last-child { border-bottom: none; }
.lbl {
  width: 72px;
  flex: 0 0 auto;
  color: #666;
  padding-top: 3px;
}
.opts {
  display: flex;
  flex-direction: column;
  gap: 6px;
}
.opts.inline {
  flex-direction: row;
  flex-wrap: wrap;
  align-items: center;
  gap: 10px;
}
.opts label {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  cursor: pointer;
}
.opts label.disabled { color: #bbb; cursor: not-allowed; }
.hint { color: #99a; font-size: 12px; }
select {
  height: 26px;
  border: 1px solid #d0d3d6;
  border-radius: 4px;
  background: #fff;
  padding: 0 6px;
}
.dlg-foot {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 12px 16px;
  border-top: 1px solid #eef0f2;
  background: #fafbfc;
}
.dlg-foot .grow { flex: 1; }
.dlg-foot button {
  height: 30px;
  padding: 0 14px;
  border: 1px solid #d0d3d6;
  border-radius: 5px;
  background: #fff;
  cursor: pointer;
  font-size: 13px;
  color: #333;
}
.dlg-foot button:hover { background: #f0f2f4; }
.dlg-foot button.ghost { border-color: transparent; color: #888; }
.dlg-foot button.primary {
  background: #21a366;
  border-color: #21a366;
  color: #fff;
}
.dlg-foot button.primary:hover { background: #1c8f59; }
</style>
