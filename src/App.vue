<script setup lang="ts">
import { ref } from 'vue'
import ExcelViewer from './components/ExcelViewer.vue'

const src = ref<File | string | undefined>(undefined)
const fileName = ref<string>('')
const dragOver = ref(false)

function onFile(file: File | undefined) {
  if (!file) return
  src.value = file
  fileName.value = file.name
}

function onInput(e: Event) {
  const input = e.target as HTMLInputElement
  onFile(input.files?.[0])
}

function onDrop(e: DragEvent) {
  dragOver.value = false
  onFile(e.dataTransfer?.files?.[0])
}

async function loadSample() {
  // 用 BASE_URL 前缀,部署到子路径(GitHub Pages)也能取到示例
  src.value = import.meta.env.BASE_URL + 'sample.xlsx'
  fileName.value = 'sample.xlsx'
}
</script>

<template>
  <div
    class="app"
    :class="{ dragging: dragOver }"
    @dragover.prevent="dragOver = true"
    @dragleave.prevent="dragOver = false"
    @drop.prevent="onDrop"
  >
    <header class="app-bar">
      <strong>OOXML Excel 预览器</strong>
      <span class="sub">Vue3 · Canvas 高保真 · 只读</span>
      <div class="grow" />
      <label class="file-btn">
        选择 .xlsx
        <input type="file" accept=".xlsx,.xlsm" @change="onInput" hidden />
      </label>
      <button class="sample-btn" @click="loadSample">加载示例</button>
    </header>

    <main class="app-body">
      <ExcelViewer :src="src" :file-name="fileName" />
      <div v-if="dragOver" class="drop-hint">松开以加载文件</div>
    </main>
  </div>
</template>

<style scoped>
.app {
  display: flex;
  flex-direction: column;
  height: 100vh;
  width: 100vw;
}
.app.dragging { outline: 3px dashed #21a366; outline-offset: -6px; }
.app-bar {
  display: flex;
  align-items: center;
  gap: 12px;
  height: 48px;
  padding: 0 16px;
  background: #21a366;
  color: #fff;
}
.app-bar .sub { font-size: 12px; opacity: 0.85; }
.app-bar .grow { flex: 1; }
.file-btn, .sample-btn {
  background: rgba(255, 255, 255, 0.18);
  color: #fff;
  border: 1px solid rgba(255, 255, 255, 0.4);
  padding: 6px 14px;
  border-radius: 5px;
  cursor: pointer;
  font-size: 13px;
}
.file-btn:hover, .sample-btn:hover { background: rgba(255, 255, 255, 0.3); }
.app-body {
  position: relative;
  flex: 1 1 auto;
  min-height: 0;
}
.drop-hint {
  position: absolute;
  inset: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(33, 163, 102, 0.08);
  font-size: 20px;
  color: #21a366;
  pointer-events: none;
}
</style>
