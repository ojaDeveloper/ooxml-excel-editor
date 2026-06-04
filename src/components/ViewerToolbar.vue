<script setup lang="ts">
import { onBeforeUnmount, ref } from 'vue'

defineProps<{
  fileName?: string
  sheetCount: number
  zoom: number
}>()
const emit = defineEmits<{
  (e: 'update:zoom', value: number): void
  (e: 'export-image'): void
  (e: 'export-pdf'): void
  (e: 'export-pdf-vector'): void
  (e: 'print'): void
  (e: 'open-settings'): void
}>()

const STEPS = [0.5, 0.75, 1, 1.25, 1.5, 2]

function setZoom(z: number) {
  emit('update:zoom', Math.min(3, Math.max(0.3, z)))
}

// 导出下拉
const menuOpen = ref(false)
function toggleMenu() {
  menuOpen.value = !menuOpen.value
}
type MenuAction = 'export-image' | 'export-pdf' | 'export-pdf-vector' | 'print' | 'open-settings'
function pick(action: MenuAction) {
  menuOpen.value = false
  if (action === 'export-image') emit('export-image')
  else if (action === 'export-pdf') emit('export-pdf')
  else if (action === 'export-pdf-vector') emit('export-pdf-vector')
  else if (action === 'print') emit('print')
  else emit('open-settings')
}
function onDocClick(e: MouseEvent) {
  if (!(e.target as HTMLElement)?.closest('.export-wrap')) menuOpen.value = false
}
if (typeof document !== 'undefined') document.addEventListener('click', onDocClick)
onBeforeUnmount(() => {
  if (typeof document !== 'undefined') document.removeEventListener('click', onDocClick)
})
</script>

<template>
  <div class="toolbar">
    <span class="file" :title="fileName">{{ fileName || '未命名工作簿' }}</span>
    <span class="meta">{{ sheetCount }} 个工作表</span>
    <div class="spacer" />

    <div class="export-wrap">
      <button class="export-btn" @click.stop="toggleMenu" title="导出 / 打印">
        导出 <span class="caret">▾</span>
      </button>
      <div v-if="menuOpen" class="menu">
        <button @click="pick('export-image')">导出为图片 (PNG)</button>
        <button @click="pick('export-pdf')">导出为 PDF (位图)</button>
        <button @click="pick('export-pdf-vector')">导出为 PDF (矢量·文字可选)</button>
        <button @click="pick('print')">打印…</button>
        <div class="sep" />
        <button @click="pick('open-settings')">导出设置…</button>
      </div>
    </div>

    <div class="zoom">
      <button @click="setZoom(zoom - 0.1)" title="缩小">−</button>
      <select :value="zoom" @change="setZoom(parseFloat(($event.target as HTMLSelectElement).value))">
        <option v-for="s in STEPS" :key="s" :value="s">{{ Math.round(s * 100) }}%</option>
        <option v-if="!STEPS.includes(zoom)" :value="zoom">{{ Math.round(zoom * 100) }}%</option>
      </select>
      <button @click="setZoom(zoom + 0.1)" title="放大">+</button>
    </div>
  </div>
</template>

<style scoped>
.toolbar {
  display: flex;
  align-items: center;
  gap: 12px;
  height: 36px;
  padding: 0 12px;
  background: #fbfbfb;
  border-bottom: 1px solid #e2e4e7;
  font-size: 13px;
  color: #444;
  flex: 0 0 auto;
}
.file {
  font-weight: 600;
  max-width: 320px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.meta { color: #888; }
.spacer { flex: 1; }

.export-wrap { position: relative; }
.export-btn {
  height: 24px;
  padding: 0 10px;
  border: 1px solid #d0d3d6;
  background: #fff;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
  color: #444;
  display: flex;
  align-items: center;
  gap: 4px;
}
.export-btn:hover { background: #f0f2f4; }
.export-btn .caret { font-size: 10px; color: #888; }
.menu {
  position: absolute;
  right: 0;
  top: 28px;
  z-index: 20;
  min-width: 160px;
  background: #fff;
  border: 1px solid #d8dbde;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.14);
  padding: 4px;
  display: flex;
  flex-direction: column;
}
.menu button {
  text-align: left;
  padding: 7px 10px;
  border: none;
  background: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
  color: #333;
  white-space: nowrap;
}
.menu button:hover { background: #f0f4ff; }
.menu .sep { height: 1px; background: #eef0f2; margin: 4px 2px; }

.zoom { display: flex; align-items: center; gap: 4px; }
.zoom button {
  width: 24px;
  height: 24px;
  border: 1px solid #d0d3d6;
  background: #fff;
  border-radius: 4px;
  cursor: pointer;
  font-size: 15px;
  line-height: 1;
}
.zoom button:hover { background: #f0f2f4; }
.zoom select {
  height: 24px;
  border: 1px solid #d0d3d6;
  border-radius: 4px;
  background: #fff;
}
</style>
