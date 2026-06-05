<script setup lang="ts">
import { ref, watch } from 'vue'
import ExcelViewer from './components/ExcelViewer.vue'
import { definePlugin } from './core/plugin'
import type { ViewerApi } from './core/plugin'
import type { PdfPageContext } from './core/export/types'
import { demoSelectEditor } from './demo-shared/demo-editor'

const src = ref<File | string | undefined>(undefined)
const fileName = ref<string>('')
const dragOver = ref(false)
const editMode = ref(false) // E0: 编辑模式闸门(默认只读)

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

// ---- 扩展 API 演示 ----
const lastEvent = ref('')
const viewerRef = ref<ViewerApi | null>(null)

// 编辑变更:记录到状态栏 + (DEV)挂 window 供 e2e 校验前后快照
function onCellChange(p: { before: { text: string }; after: { text: string }; source: string }) {
  lastEvent.value = `[${p.source}] R?C? "${p.before.text}" → "${p.after.text}"`
  if (import.meta.env.DEV) (window as unknown as { __lastCellChange?: unknown }).__lastCellChange = p
}
// E3.5: 列宽/行高 + 脏状态变更(DEV 挂 window 供 e2e 校验)
function onDimChange(p: { axis: string; index: number; before: number; after: number; source: string }) {
  lastEvent.value = `[${p.source}] ${p.axis}${p.index} ${Math.round(p.before)}→${Math.round(p.after)}px`
  if (import.meta.env.DEV) (window as unknown as { __lastDimChange?: unknown }).__lastDimChange = p
}
function onDirtyChange(p: { dirty: boolean }) {
  if (import.meta.env.DEV) (window as unknown as { __lastDirtyChange?: unknown }).__lastDirtyChange = p
}
function onImageChange(p: unknown) {
  if (import.meta.env.DEV) (window as unknown as { __lastImageChange?: unknown }).__lastImageChange = p
}
function onStructChange(p: unknown) {
  if (import.meta.env.DEV) (window as unknown as { __lastStructChange?: unknown }).__lastStructChange = p
}
// E7: 在选区上方插入一行 / 删除选区所在行
function insertRowAtSel() {
  const v = viewerRef.value
  const sel = v?.getSelection()
  if (v && sel) v.insertRows(sel.top, 1)
}
function deleteRowAtSel() {
  const v = viewerRef.value
  const sel = v?.getSelection()
  if (v && sel) v.deleteRows(sel.top, sel.bottom - sel.top + 1)
}
// E5: 给当前选区加粗(样式编辑演示)
function boldSelection() {
  const v = viewerRef.value
  const sel = v?.getSelection()
  if (v && sel) v.setStyle(sel, { font: { bold: true } })
}
// G1: 合并 / 拆分当前选区
function mergeSelection() {
  const v = viewerRef.value
  const sel = v?.getSelection()
  if (v && sel) v.mergeCells(sel)
}
function unmergeSelection() {
  const v = viewerRef.value
  const sel = v?.getSelection()
  if (v && sel) v.unmergeCells(sel)
}

// 开发环境把命令式 API 挂到 window,便于 e2e 计算 canvas 上的几何(如筛选按钮位置)
if (import.meta.env.DEV) {
  watch(viewerRef, (v) => {
    ;(window as unknown as { __excelViewer?: ViewerApi | null }).__excelViewer = v
  })
}

// 演示 beforeRenderPage 扩展钩子: 每页右下角页码 + 居中淡水印
async function exportPdfWithWatermark() {
  const viewer = viewerRef.value
  if (!viewer) return
  try {
    await viewer.downloadPdf({
      target: 'all',
      beforeRenderPage: (ctx: PdfPageContext) => {
        const { doc, pageIndex, pageCount, pageWidth, pageHeight, margin, sheetName } = ctx
        // 页脚: 表名 + 页码
        doc.setFontSize(9)
        doc.setTextColor(120)
        doc.text(`${sheetName}`, margin.left, pageHeight - 5)
        doc.text(`第 ${pageIndex + 1} / ${pageCount} 页`, pageWidth - margin.right, pageHeight - 5, { align: 'right' })
        // 水印: 居中旋转大字
        doc.setFontSize(56)
        doc.setTextColor(230)
        doc.text('PREVIEW', pageWidth / 2, pageHeight / 2, { align: 'center', angle: 30 })
      },
    })
    lastEvent.value = '已导出 PDF(全部表 + 页码 + 水印)'
  } catch (e) {
    lastEvent.value = '导出失败: ' + (e as Error).message
  }
}

// 演示数据读取 API: 取当前表为 JSON(示例表头在第 1 行)→ 复制到剪贴板 + toast
function showSheetJSON() {
  const viewer = viewerRef.value
  if (!viewer) return
  const json = viewer.getSheetJSON({ headerRow: 1 })
  navigator.clipboard?.writeText(JSON.stringify(json, null, 2)).catch(() => {})
  lastEvent.value = `[数据] ${json.length} 行已复制为 JSON · 首行: ${JSON.stringify(json[0] ?? {})}`.slice(0, 140)
}

// 示例插件: 负数标红 + 单击 toast + 贡献一个工具栏按钮(definePlugin 打包 cellStyle/events/toolbar)
const negativesPlugin = definePlugin({
  name: 'demo-highlight-negatives',
  cellStyle: (cell) =>
    typeof cell.raw === 'number' && cell.raw < 0 ? { font: { color: '#d00', bold: true } } : undefined,
  events: {
    'cell-click': (p) => (lastEvent.value = `[插件] 点击 R${p.row + 1}C${p.col + 1}: ${p.text}`),
  },
  toolbar: [
    {
      id: 'demo-info',
      icon: '🔔',
      label: '插件按钮',
      title: '演示插件贡献的工具栏项',
      onClick: (viewer) =>
        (lastEvent.value = `[插件工具栏] 当前 ${viewer.getWorkbook()?.sheets.length ?? 0} 个工作表`),
    },
  ],
  // overlay 返回 DOM(框架无关,跟 React 壳同一份写法):A1(row0,col0)叠个徽标
  overlay: ({ rectOf }) => {
    const r = rectOf(0, 0)
    if (!r) return null
    const el = document.createElement('div')
    el.className = 'plugin-badge'
    el.textContent = '🔌'
    Object.assign(el.style, { position: 'absolute', left: r.x + r.w - 14 + 'px', top: r.y - 2 + 'px', fontSize: '12px', pointerEvents: 'none' })
    return el
  },
})
const plugins = [negativesPlugin]
type Rect = { x: number; y: number; w: number; h: number } | null
// overlay slot: 在 B3(row2,col1)叠一个徽标,随滚动跟随;tick 变化触发重算
function badgeStyle(rectOf: (r: number, c: number) => Rect, _tick: number) {
  const r = rectOf(2, 1)
  if (!r) return { display: 'none' }
  return {
    position: 'absolute' as const,
    left: r.x + r.w - 20 + 'px',
    top: r.y + 1 + 'px',
    display: r.x + r.w < 0 || r.y < 0 ? 'none' : 'block',
  }
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
      <button v-if="src" class="sample-btn" @click="exportPdfWithWatermark" title="演示 beforeRenderPage 钩子">
        PDF(页码+水印)
      </button>
      <button v-if="src" class="sample-btn" @click="showSheetJSON" title="演示数据读取 API getSheetJSON">
        数据→JSON
      </button>
      <label v-if="src" class="edit-toggle" title="开启编辑模式(E0:闸门)">
        <input type="checkbox" v-model="editMode" /> 编辑模式
      </label>
      <button v-if="src && editMode" class="sample-btn" @click="boldSelection" title="给选区加粗(E5:样式编辑)">
        B 加粗选区
      </button>
      <button v-if="src && editMode" class="sample-btn" @click="mergeSelection" title="合并选区(G1)">合并</button>
      <button v-if="src && editMode" class="sample-btn" @click="unmergeSelection" title="拆分选区(G1)">拆分</button>
      <button v-if="src && editMode" class="sample-btn" @click="insertRowAtSel" title="选区上方插入行(E7)">＋行</button>
      <button v-if="src && editMode" class="sample-btn" @click="deleteRowAtSel" title="删除选区行(E7)">－行</button>
      <button v-if="src" class="sample-btn" @click="viewerRef?.downloadXlsx()" title="导出 .xlsx(E8:从模型重建)">↓XLSX</button>
      <button v-if="src" class="sample-btn" @click="viewerRef?.downloadCsv()" title="导出 .csv(E8)">↓CSV</button>
      <button v-if="src" class="sample-btn" @click="viewerRef?.downloadJson()" title="导出 .json(E8)">↓JSON</button>
    </header>

    <main class="app-body">
      <ExcelViewer
        ref="viewerRef"
        :src="src"
        :file-name="fileName"
        :plugins="plugins"
        :editable="editMode"
        :recalc="editMode"
        :read-only-ranges="[{ top: 1, left: 0, bottom: 1, right: 4 }]"
        :editor="demoSelectEditor"
        :toolbar="['find', 'filter', 'clear-filter', 'separator', 'copy', 'freeze', 'separator', 'zoom', 'export']"
        @selection-change="(s) => (lastEvent = `选区 ${s.range.top + 1},${s.range.left + 1} → ${s.range.bottom + 1},${s.range.right + 1}`)"
        @cell-change="onCellChange"
        @dim-change="onDimChange"
        @dirty-change="onDirtyChange"
        @image-change="onImageChange"
        @struct-change="onStructChange"
      >
        <!-- 分层 UI 演示: B3 上叠一个可点徽标,随滚动跟随 -->
        <template #overlay="{ rectOf, tick }">
          <div
            class="demo-badge"
            :style="badgeStyle(rectOf, tick)"
            title="overlay slot 演示(锚在 B3)"
            @click="lastEvent = '点了叠加层徽标'"
          >
            ★
          </div>
        </template>
      </ExcelViewer>
      <div v-if="lastEvent" class="event-toast">{{ lastEvent }}</div>
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
.demo-badge {
  width: 18px;
  height: 18px;
  border-radius: 50%;
  background: #ff6b35;
  color: #fff;
  font-size: 12px;
  line-height: 18px;
  text-align: center;
  cursor: pointer;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.3);
}
.event-toast {
  position: absolute;
  left: 12px;
  bottom: 12px;
  background: rgba(0, 0, 0, 0.78);
  color: #fff;
  font-size: 12px;
  padding: 5px 10px;
  border-radius: 5px;
  pointer-events: none;
  z-index: 10;
}
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
