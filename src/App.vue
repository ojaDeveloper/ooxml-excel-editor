<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
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
  selTick.value++ // 内容/样式变 → 颜色回显重算
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
// WPS 内嵌图 ⇄ 浮动图互转
function embedAll() {
  const n = viewerRef.value?.convertAllImagesToCells() ?? 0 // 整表就近嵌入
  if (!n) alert('没有可嵌入的浮动图')
}
function cellToFloat() {
  const v = viewerRef.value
  const sel = v?.getSelection()
  if (v && sel) v.convertCellImageToFloat(sel.top, sel.left)
}
const cellImageFit = ref<'fill' | 'contain' | 'cover'>('contain')
// 背景色 / 字体色:回显当前活动格 + 改选区(WPS 风格)
const selTick = ref(0)
const activeFill = computed(() => {
  void selTick.value
  return viewerRef.value?.getActiveFillColor() ?? '#FFFFFF'
})
const activeFont = computed(() => {
  void selTick.value
  return viewerRef.value?.getActiveFontColor() ?? '#000000'
})
function setFill(e: Event) {
  viewerRef.value?.setSelectionFill((e.target as HTMLInputElement).value)
  selTick.value++
}
function setFont(e: Event) {
  viewerRef.value?.setSelectionFontColor((e.target as HTMLInputElement).value)
  selTick.value++
}
function clearFill() {
  viewerRef.value?.setSelectionFill(null)
  selTick.value++
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

// ---------------- demo 顶栏:按钮太多自动收进「⋯ 更多」 ----------------
// 把可溢出的演示按钮抽成数据数组,做隐藏测量 + 可见/溢出切分。本逻辑只在 demo 里,
// 跟组件内置的 ActionToolbar 各管各的:demo 的演示按钮永远不进 :toolbar(那是组件的领地)。
interface DemoItem {
  id: string
  /** 'btn' 普通按钮 / 'color' 取色器(label+input[type=color]) / 'select' 下拉 */
  type: 'btn' | 'color' | 'select'
  label: string
  title?: string
  when?: () => boolean
  onClick?: () => void
  /** color: 当前回显色 + 改色 */
  getColor?: () => string
  onColorInput?: (e: Event) => void
  /** select: 选项列表 + v-model */
  selectModel?: () => string
  selectOptions?: { value: string; label: string }[]
  onSelect?: (v: string) => void
}
const demoBarItems = computed<DemoItem[]>(() => {
  if (!src.value) return []
  void selTick.value // 颜色回显跟 selTick 走
  const arr: DemoItem[] = [
    { id: 'pdf-watermark', type: 'btn', label: 'PDF(页码+水印)', title: '演示 beforeRenderPage 钩子', onClick: () => void exportPdfWithWatermark() },
    { id: 'sheet-json', type: 'btn', label: '数据→JSON', title: '演示数据读取 API getSheetJSON', onClick: showSheetJSON },
    { id: 'fit', type: 'select', label: '贴合', title: 'WPS 内嵌图贴合方式',
      selectModel: () => cellImageFit.value,
      selectOptions: [
        { value: 'contain', label: 'contain 等比(同 WPS)' },
        { value: 'fill', label: 'fill 铺满' },
        { value: 'cover', label: 'cover 裁剪' },
      ],
      onSelect: (v) => (cellImageFit.value = v as 'fill' | 'contain' | 'cover'),
    },
    { id: 'dl-xlsx', type: 'btn', label: '↓XLSX', title: '导出 .xlsx(E8:从模型重建)', onClick: () => void viewerRef.value?.downloadXlsx() },
    { id: 'dl-csv', type: 'btn', label: '↓CSV', title: '导出 .csv(E8)', onClick: () => viewerRef.value?.downloadCsv() },
    { id: 'dl-json', type: 'btn', label: '↓JSON', title: '导出 .json(E8)', onClick: () => viewerRef.value?.downloadJson() },
  ]
  if (editMode.value) {
    arr.unshift(
      { id: 'bold', type: 'btn', label: 'B 加粗选区', title: '给选区加粗(E5)', onClick: boldSelection },
      { id: 'merge', type: 'btn', label: '合并', title: '合并选区(G1)', onClick: mergeSelection },
      { id: 'unmerge', type: 'btn', label: '拆分', title: '拆分选区(G1)', onClick: unmergeSelection },
      { id: 'fill', type: 'color', label: '背景', title: '背景填充色(回显 + 改选区,WPS 风格)', getColor: () => activeFill.value, onColorInput: setFill },
      { id: 'font', type: 'color', label: '字体', title: '字体颜色(回显 + 改选区)', getColor: () => activeFont.value, onColorInput: setFont },
      { id: 'clear-fill', type: 'btn', label: '清除填充', title: '清除背景填充(还原无填充/白)', onClick: clearFill },
      { id: 'embed-all', type: 'btn', label: '整表嵌入', title: '整表浮动图就近嵌入(WPS 浮动→嵌入/DISPIMG)', onClick: embedAll },
      { id: 'cell-to-float', type: 'btn', label: '格→图', title: '把选中格的内嵌图拎成浮动图(WPS 嵌入→浮动)', onClick: cellToFloat },
      { id: 'ins-row', type: 'btn', label: '＋行', title: '选区上方插入行(E7)', onClick: insertRowAtSel },
      { id: 'del-row', type: 'btn', label: '－行', title: '删除选区行(E7)', onClick: deleteRowAtSel },
    )
  }
  return arr
})

// 测量 + 计算可见数(放不下的进「更多」popover)
const demoBarEl = ref<HTMLElement | null>(null)
const demoMeasureEl = ref<HTMLElement | null>(null)
const demoItemWidths = ref<number[]>([])
const demoBarContentW = ref(0)
const demoMoreOpen = ref(false)
const DEMO_MORE_W = 64
const DEMO_GAP = 6
function demoRemeasure() {
  const el = demoMeasureEl.value
  if (!el) return
  demoItemWidths.value = Array.from(el.children).map((c) => (c as HTMLElement).offsetWidth)
  const bar = demoBarEl.value
  if (!bar) return
  // 剩余可用宽 = bar 总宽 - 固定区(strong/sub/grow/file/sample/edit)实际占用,后者是 bar 的直接子节点除测量+可见 row + more 之外的部分
  const fixed = Array.from(bar.children).find((c) => (c as HTMLElement).classList.contains('app-bar-fixed')) as HTMLElement | null
  const fixedW = fixed ? fixed.getBoundingClientRect().width : 0
  demoBarContentW.value = Math.max(0, bar.clientWidth - fixedW - 24 /* 边距 */)
}
const demoVisibleCount = computed(() => {
  const cw = demoBarContentW.value
  const w = demoItemWidths.value
  const items = demoBarItems.value
  if (!cw || w.length !== items.length) return items.length
  let sum = 0
  let fitsAll = true
  for (let i = 0; i < items.length; i++) {
    sum += w[i] + DEMO_GAP
    if (sum > cw) { fitsAll = false; break }
  }
  if (fitsAll) return items.length
  let s = DEMO_MORE_W
  let n = 0
  for (let i = 0; i < items.length; i++) {
    s += w[i] + DEMO_GAP
    if (s > cw) break
    n++
  }
  return Math.max(0, n)
})
const demoVisibleItems = computed(() => demoBarItems.value.slice(0, demoVisibleCount.value))
const demoOverflowItems = computed(() => demoBarItems.value.slice(demoVisibleCount.value))
let demoRo: ResizeObserver | null = null
function onDemoBarPointer(e: MouseEvent) {
  if (demoBarEl.value && !demoBarEl.value.contains(e.target as Node)) demoMoreOpen.value = false
}
onMounted(() => {
  nextTick(demoRemeasure)
  demoRo = new ResizeObserver(() => demoRemeasure())
  if (demoBarEl.value) demoRo.observe(demoBarEl.value)
  document.addEventListener('mousedown', onDemoBarPointer)
})
onBeforeUnmount(() => {
  demoRo?.disconnect()
  document.removeEventListener('mousedown', onDemoBarPointer)
})
watch(demoBarItems, () => nextTick(demoRemeasure))

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
    <!-- demo 顶栏:固定区(标题/选 xlsx/加载示例/编辑模式)+ 演示按钮区(自动溢出收进「⋯ 更多」) -->
    <header class="app-bar" ref="demoBarEl">
      <!-- 固定区:永不溢出(基础框架控件) -->
      <div class="app-bar-fixed">
        <strong>OOXML Excel 预览器</strong>
        <span class="sub">Vue3 · Canvas 高保真</span>
        <label class="file-btn">
          选择 .xlsx
          <input type="file" accept=".xlsx,.xlsm" @change="onInput" hidden />
        </label>
        <button class="sample-btn" @click="loadSample">加载示例</button>
        <label v-if="src" class="edit-toggle" title="开启编辑模式(E0:闸门)">
          <input type="checkbox" v-model="editMode" /> 编辑模式
        </label>
      </div>
      <div class="grow" />

      <!-- 隐藏测量行:为每项算实际宽度,跟可见行的样式一致 -->
      <div class="app-bar-measure" ref="demoMeasureEl" aria-hidden="true">
        <template v-for="it in demoBarItems" :key="'m' + it.id">
          <button v-if="it.type === 'btn'" class="sample-btn">{{ it.label }}</button>
          <label v-else-if="it.type === 'color'" class="sample-label">{{ it.label }}<input type="color" /></label>
          <label v-else class="sample-label">{{ it.label }}<select><option v-for="o in it.selectOptions" :key="o.value" :value="o.value">{{ o.label }}</option></select></label>
        </template>
      </div>

      <!-- 可见演示按钮 -->
      <template v-for="it in demoVisibleItems" :key="it.id">
        <button v-if="it.type === 'btn'" class="sample-btn" :title="it.title" @click="it.onClick?.()">{{ it.label }}</button>
        <label v-else-if="it.type === 'color'" class="sample-label" :title="it.title">{{ it.label }}<input type="color" :value="it.getColor?.()" @input="it.onColorInput?.($event)" /></label>
        <label v-else class="sample-label" :title="it.title">{{ it.label }}<select :value="it.selectModel?.()" @change="it.onSelect?.(($event.target as HTMLSelectElement).value)"><option v-for="o in it.selectOptions" :key="o.value" :value="o.value">{{ o.label }}</option></select></label>
      </template>

      <!-- 更多溢出 popover -->
      <div v-if="demoOverflowItems.length" class="more-wrap">
        <button class="sample-btn more-btn" :class="{ open: demoMoreOpen }" title="更多" @click="demoMoreOpen = !demoMoreOpen">⋯ 更多</button>
        <div v-if="demoMoreOpen" class="more-pop">
          <template v-for="it in demoOverflowItems" :key="'o' + it.id">
            <button v-if="it.type === 'btn'" class="more-row" :title="it.title" @click="it.onClick?.(); demoMoreOpen = false">{{ it.label }}</button>
            <label v-else-if="it.type === 'color'" class="more-row" :title="it.title">{{ it.label }}<input type="color" :value="it.getColor?.()" @input="it.onColorInput?.($event)" /></label>
            <label v-else class="more-row" :title="it.title">{{ it.label }}<select :value="it.selectModel?.()" @change="it.onSelect?.(($event.target as HTMLSelectElement).value); demoMoreOpen = false"><option v-for="o in it.selectOptions" :key="o.value" :value="o.value">{{ o.label }}</option></select></label>
          </template>
        </div>
      </div>
    </header>

    <main class="app-body">
      <ExcelViewer
        ref="viewerRef"
        :src="src"
        :file-name="fileName"
        :plugins="plugins"
        :cell-image-fit="cellImageFit"
        :editable="editMode"
        :recalc="editMode"
        :read-only-ranges="[{ top: 1, left: 0, bottom: 1, right: 4 }]"
        :editor="demoSelectEditor"
        :toolbar="['find', 'filter', 'clear-filter', 'separator', 'copy', 'wrap-text', 'freeze', 'separator', 'zoom', 'export']"
        @selection-change="(s) => { lastEvent = `选区 ${s.range.top + 1},${s.range.left + 1} → ${s.range.bottom + 1},${s.range.right + 1}`; selTick++ }"
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
  position: relative;
  display: flex;
  align-items: center;
  gap: 8px;
  height: 48px;
  padding: 0 16px;
  background: #21a366;
  color: #fff;
  overflow: visible; /* 让「更多」popover 能溢出到栏外 */
}
.app-bar-fixed {
  display: flex;
  align-items: center;
  gap: 12px;
  flex: 0 0 auto;
}
.app-bar .sub { font-size: 12px; opacity: 0.85; }
.app-bar .grow { flex: 1 1 0; min-width: 0; }
.file-btn, .sample-btn {
  background: rgba(255, 255, 255, 0.18);
  color: #fff;
  border: 1px solid rgba(255, 255, 255, 0.4);
  padding: 6px 14px;
  border-radius: 5px;
  cursor: pointer;
  font-size: 13px;
  white-space: nowrap;
}
.file-btn:hover, .sample-btn:hover { background: rgba(255, 255, 255, 0.3); }
.sample-label {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: 13px;
  white-space: nowrap;
}
.sample-label input[type=color], .sample-label select { vertical-align: middle; }
/* 隐藏测量行:占位但不显示 + 不响应交互 */
.app-bar-measure {
  position: absolute;
  left: 0; top: 0;
  visibility: hidden;
  pointer-events: none;
  height: 0;
  overflow: hidden;
  display: flex;
  gap: 6px;
}
.more-wrap { position: relative; }
.more-btn { display: inline-flex; align-items: center; gap: 4px; }
.more-btn.open { background: rgba(255, 255, 255, 0.32); }
.more-pop {
  position: absolute;
  top: calc(100% + 4px);
  right: 0;
  z-index: 30;
  min-width: 200px;
  background: #fff;
  color: #1f2329;
  border: 1px solid #d8dbde;
  border-radius: 6px;
  box-shadow: 0 6px 18px rgba(0, 0, 0, 0.18);
  padding: 4px;
  display: flex;
  flex-direction: column;
  align-items: stretch;
  gap: 2px;
}
.more-row {
  display: flex;
  align-items: center;
  gap: 8px;
  text-align: left;
  background: none;
  border: none;
  font: inherit;
  font-size: 13px;
  color: #1f2329;
  padding: 7px 10px;
  border-radius: 4px;
  cursor: pointer;
}
.more-row:hover { background: #eef3fe; }
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
