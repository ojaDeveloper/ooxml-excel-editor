<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import ExcelViewer from './components/ExcelViewer.vue'
import { definePlugin } from './core/plugin'
import type { ViewerApi } from './core/plugin'
import type { EditableTarget } from './core/edit/types'
import type { PdfPageContext } from './core/export/types'
import { demoSelectEditor } from './demo-shared/demo-editor'

const src = ref<File | string | undefined>(undefined)
const fileName = ref<string>('')
const dragOver = ref(false)
const editMode = ref(false) // E0: 编辑模式闸门(默认只读)

// JSON 数据源 + 模板样式 overlay(P3 重设计 2026-06-08)演示
// 模板的语义改了 —— 不再是"占位符 + 锚点表"那套数据替换;现在是"样式捐赠者":数据在 A1 自然位置,
// 模板贡献 styling (字体/边框/列宽/合并/freeze),模板的文字内容全部丢弃.
// 因此 demo 不再传 placeholders/anchors,直接传 JSON items 即可;工具栏「模板 ▾」导入 .xlsx 即生效.
const jsonItems = ref<Array<Record<string, unknown>> | null>(null)

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
  jsonItems.value = null // 切回 .xlsx 模式
}
async function loadJsonSample() {
  // JSON 数据源演示:不传 src,只给 :workbook(对象数组)→ 默认渲染(数据在 A1 起自然位置)
  // 然后点工具栏「模板 ▾ → 导入 .xlsx」可套用模板样式 (模板的文字内容会被丢弃,只保留 styling)
  src.value = undefined
  fileName.value = '订单数据'
  jsonItems.value = [
    { name: '笔记本电脑', price: 5999, qty: 1, amount: 5999, note: '商务款' },
    { name: '机械键盘', price: 399, qty: 2, amount: 798, note: '青轴' },
    { name: '显示器', price: 1299, qty: 2, amount: 2598, note: '27寸 2K' },
    { name: '鼠标', price: 89, qty: 5, amount: 445, note: '无线' },
    { name: '耳机', price: 599, qty: 3, amount: 1797, note: '降噪' },
  ]
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
// Phase A 2026-06-08: 权限拒绝事件 → 红色 toast 提示, 让用户能感知"为啥点了没反应"
function onPermissionDenied(p: { reason: string; cells: Array<{ row: number; col: number }>; dims?: { axis: string; indices: number[] }; message?: string }) {
  const reasonLabel = ({ paste: '粘贴', merge: '合并', unmerge: '拆分', 'image-place': '放图', 'image-convert': '图片转换', dimension: '改尺寸', other: '操作' } as Record<string, string>)[p.reason] ?? p.reason
  const detail = p.dims ? `${p.dims.indices.length} 个 ${p.dims.axis}` : `${p.cells.length} 个格`
  lastEvent.value = `🚫 [${reasonLabel}] 拒绝: ${detail} 在白名单外 — ${p.message ?? ''}`
  if (import.meta.env.DEV) (window as unknown as { __lastPermissionDenied?: unknown }).__lastPermissionDenied = p
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

// ---- 「设置可编辑单元格」演示(白名单 API,2026-06-08) ----
// 默认 undefined = 白名单未启用(老行为:editable=true 时全可编辑);
// applied = [] 或 [...targets] = 白名单生效(只 targets 可编辑)
const editableTargetsApplied = ref<EditableTarget[] | undefined>(undefined)
// Phase C 2026-06-08: 高亮只读 toggle - 把只读格套浅灰底, 让用户一眼区分白名单内外
const highlightReadOnly = ref(false)
const editTargetsDialogOpen = ref(false)
// 弹窗里的临时选区: 用 "r:c" 字符串集合记录(可独立勾选不相邻格)
const editTargetsDraft = ref<Set<string>>(new Set())
// 弹窗里"行/列整选"勾(独立于单格勾)
const editTargetsRowDraft = ref<Set<number>>(new Set())
const editTargetsColDraft = ref<Set<number>>(new Set())

const EDIT_DIALOG_ROWS = 12
const EDIT_DIALOG_COLS = 8

function openEditTargetsDialog() {
  // 用上一次应用的状态回灌, 没就空开
  editTargetsDraft.value = new Set()
  editTargetsRowDraft.value = new Set()
  editTargetsColDraft.value = new Set()
  for (const t of editableTargetsApplied.value ?? []) {
    if ('top' in t) continue // 矩形跳过(demo 只用单格/整行/整列)
    if ('row' in t && 'col' in t && typeof t.col === 'number') {
      editTargetsDraft.value.add(`${t.row}:${t.col}`)
    } else if ('row' in t && typeof t.row === 'number') {
      editTargetsRowDraft.value.add(t.row)
    } else if ('col' in t && typeof t.col === 'number') {
      editTargetsColDraft.value.add(t.col)
    }
  }
  editTargetsDialogOpen.value = true
}
function toggleEditTargetCell(r: number, c: number) {
  const k = `${r}:${c}`
  if (editTargetsDraft.value.has(k)) editTargetsDraft.value.delete(k)
  else editTargetsDraft.value.add(k)
  editTargetsDraft.value = new Set(editTargetsDraft.value) // trigger reactivity
}
function toggleEditTargetRow(r: number) {
  if (editTargetsRowDraft.value.has(r)) editTargetsRowDraft.value.delete(r)
  else editTargetsRowDraft.value.add(r)
  editTargetsRowDraft.value = new Set(editTargetsRowDraft.value)
}
function toggleEditTargetCol(c: number) {
  if (editTargetsColDraft.value.has(c)) editTargetsColDraft.value.delete(c)
  else editTargetsColDraft.value.add(c)
  editTargetsColDraft.value = new Set(editTargetsColDraft.value)
}
function isCellInDraft(r: number, c: number): boolean {
  return editTargetsDraft.value.has(`${r}:${c}`) ||
    editTargetsRowDraft.value.has(r) ||
    editTargetsColDraft.value.has(c)
}
function applyEditTargets() {
  const arr: EditableTarget[] = []
  // 整行 / 整列 优先(更省内存 + 范围更直观)
  for (const r of editTargetsRowDraft.value) arr.push({ row: r })
  for (const c of editTargetsColDraft.value) arr.push({ col: c })
  // 单格(排除已经被整行/整列覆盖的, 避免重复)
  for (const k of editTargetsDraft.value) {
    const [r, c] = k.split(':').map(Number)
    if (editTargetsRowDraft.value.has(r) || editTargetsColDraft.value.has(c)) continue
    arr.push({ row: r, col: c })
  }
  editableTargetsApplied.value = arr
  editTargetsDialogOpen.value = false
  lastEvent.value = `[白名单] ${arr.length} 项 target 已应用; 其它格只读`
}
function clearEditTargets() {
  editableTargetsApplied.value = undefined
  editTargetsDialogOpen.value = false
  lastEvent.value = '[白名单] 已关闭, 恢复默认 (全可编辑)'
}
/** 弹窗里某格的预览文字(显示用), 取自当前工作簿;无则空 */
function previewCellText(r: number, c: number): string {
  const v = viewerRef.value?.getCellText(r, c) ?? ''
  return v.length > 6 ? v.slice(0, 6) + '…' : v
}
function colLetter(c: number): string {
  let s = ''
  let n = c
  while (true) { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; if (n < 0) break }
  return s
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
function jumpToLastRow() {
  const viewer = viewerRef.value
  const wb = viewer?.getWorkbook()
  if (!viewer || !wb) return
  const sheet = wb.sheets[viewer.getActiveSheet()]
  const row = Math.max(0, sheet.dimension.rows - 1)
  viewer.scrollToCell(row, 0, { select: true })
  lastEvent.value = `已跳到末行 A${row + 1}`
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
    { id: 'jump-last-row', type: 'btn', label: '跳到末行', title: '演示 scrollToCell(row,col,{select:true}) 导航 API', onClick: jumpToLastRow },
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
      { id: 'edit-targets', type: 'btn',
        label: editableTargetsApplied.value ? `可编辑 (${editableTargetsApplied.value.length})` : '设置可编辑',
        title: '白名单模式: 弹窗里点选要可编辑的格 / 行 / 列, 应用后只这些可编辑 (其它一律只读)',
        onClick: openEditTargetsDialog },
      { id: 'highlight-readonly', type: 'btn',
        label: highlightReadOnly.value ? '✓ 高亮只读' : '高亮只读',
        title: '把只读格套浅灰底 (内置 readOnlyCellStyle=true). 跟「设置可编辑」白名单配合, 一眼看出哪些格可编辑',
        onClick: () => (highlightReadOnly.value = !highlightReadOnly.value) },
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
        <button class="sample-btn" @click="loadJsonSample" title="加载一个 JSON 数据源演示;然后用工具栏「模板 ▾」导入 public/template-sample.xlsx 看模板效果">JSON 示例</button>
        <label v-if="src || jsonItems" class="edit-toggle" title="开启编辑模式(E0:闸门)">
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
        :workbook="jsonItems ?? undefined"
        :file-name="fileName"
        :plugins="plugins"
        :cell-image-fit="cellImageFit"
        :editable="editMode"
        :pivot-table="true"
        :recalc="editMode"
        :read-only-ranges="[{ top: 1, left: 0, bottom: 1, right: 4 }]"
        :editable-targets="editableTargetsApplied"
        :read-only-cell-style="highlightReadOnly"
        :editor="demoSelectEditor"
        :toolbar="['find', 'filter', 'sort', 'clear-filter', 'separator', 'copy', 'pivot-table', 'wrap-text', 'image-tools', 'freeze', 'separator', 'template', 'separator', 'zoom', 'export']"
        @selection-change="(s) => { lastEvent = `选区 ${s.range.top + 1},${s.range.left + 1} → ${s.range.bottom + 1},${s.range.right + 1}`; selTick++ }"
        @cell-change="onCellChange"
        @dim-change="onDimChange"
        @dirty-change="onDirtyChange"
        @image-change="onImageChange"
        @struct-change="onStructChange"
        @permission-denied="onPermissionDenied"
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

      <!-- 「设置可编辑」对话框: 网格化点选 + 行/列整选 + 应用 (演示 editableTargets 白名单 API) -->
      <div v-if="editTargetsDialogOpen" class="edit-targets-overlay" @click.self="editTargetsDialogOpen = false">
        <div class="edit-targets-dialog">
          <header>
            <h3>设置可编辑单元格 (白名单)</h3>
            <p class="hint">
              点击单元格 = 该格可编辑;点击列标题 (A/B/C…) = 整列可编辑;点击行号 = 整行可编辑.
              应用后,只有勾选的位置可编辑,其它全部只读. 关闭白名单 = 恢复默认 (整表可编辑).
            </p>
          </header>
          <div class="edit-targets-grid">
            <table>
              <thead>
                <tr>
                  <th class="corner">#</th>
                  <th
                    v-for="c in EDIT_DIALOG_COLS"
                    :key="'h' + c"
                    :class="{ picked: editTargetsColDraft.has(c - 1) }"
                    @click="toggleEditTargetCol(c - 1)"
                    :title="`整列 ${colLetter(c - 1)} 可编辑`"
                  >{{ colLetter(c - 1) }}</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="r in EDIT_DIALOG_ROWS" :key="'r' + r">
                  <th
                    :class="{ picked: editTargetsRowDraft.has(r - 1) }"
                    @click="toggleEditTargetRow(r - 1)"
                    :title="`整行 ${r} 可编辑`"
                  >{{ r }}</th>
                  <td
                    v-for="c in EDIT_DIALOG_COLS"
                    :key="'c' + r + ',' + c"
                    :class="{ picked: isCellInDraft(r - 1, c - 1), 'row-col-hit': editTargetsRowDraft.has(r - 1) || editTargetsColDraft.has(c - 1) }"
                    @click="toggleEditTargetCell(r - 1, c - 1)"
                    :title="`R${r}C${c} (${colLetter(c - 1)}${r}) 切换`"
                  >{{ previewCellText(r - 1, c - 1) || '·' }}</td>
                </tr>
              </tbody>
            </table>
          </div>
          <footer>
            <span class="count-hint">
              已选:{{ editTargetsDraft.size }} 单格 / {{ editTargetsRowDraft.size }} 整行 / {{ editTargetsColDraft.size }} 整列
            </span>
            <button class="dlg-btn ghost" @click="editTargetsDialogOpen = false">取消</button>
            <button class="dlg-btn ghost" @click="clearEditTargets" title="移除白名单, 恢复默认 (全可编辑)">关闭白名单</button>
            <button class="dlg-btn primary" @click="applyEditTargets">应用</button>
          </footer>
        </div>
      </div>
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

/* ---- 「设置可编辑」对话框 ---- */
.edit-targets-overlay {
  position: absolute;
  inset: 0;
  background: rgba(0, 0, 0, 0.5);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 50;
}
.edit-targets-dialog {
  background: #fff;
  border-radius: 10px;
  width: min(640px, calc(100vw - 32px));
  max-height: calc(100vh - 64px);
  display: flex;
  flex-direction: column;
  box-shadow: 0 12px 32px rgba(0, 0, 0, 0.25);
  overflow: hidden;
}
.edit-targets-dialog header {
  padding: 14px 18px 8px;
  border-bottom: 1px solid #eef0f3;
}
.edit-targets-dialog h3 { margin: 0 0 6px; font-size: 15px; color: #1f2329; }
.edit-targets-dialog .hint { margin: 0; font-size: 12px; color: #707481; line-height: 1.6; }
.edit-targets-grid {
  padding: 12px 18px;
  overflow: auto;
  flex: 1 1 auto;
}
.edit-targets-grid table { border-collapse: collapse; font-size: 12px; }
.edit-targets-grid th,
.edit-targets-grid td {
  border: 1px solid #dce0e6;
  padding: 4px 6px;
  min-width: 44px;
  height: 26px;
  text-align: center;
  cursor: pointer;
  color: #1f2329;
  background: #fff;
  user-select: none;
}
.edit-targets-grid thead th,
.edit-targets-grid tbody th {
  background: #f5f7fa;
  color: #707481;
  font-weight: 600;
}
.edit-targets-grid tbody td:hover,
.edit-targets-grid thead th:hover,
.edit-targets-grid tbody th:hover { background: #eef3fe; }
.edit-targets-grid tbody td.picked,
.edit-targets-grid thead th.picked,
.edit-targets-grid tbody th.picked {
  background: #d6f0e0;
  color: #146c2e;
  font-weight: 600;
}
.edit-targets-grid tbody td.row-col-hit {
  background: linear-gradient(135deg, #eef9f1 25%, transparent 25%, transparent 50%, #eef9f1 50%, #eef9f1 75%, transparent 75%);
  background-size: 8px 8px;
}
.edit-targets-grid .corner { background: #ebeef2 !important; cursor: default; }
.edit-targets-dialog footer {
  padding: 10px 18px;
  border-top: 1px solid #eef0f3;
  display: flex;
  align-items: center;
  gap: 8px;
  background: #fafbfc;
}
.edit-targets-dialog footer .count-hint {
  flex: 1 1 auto;
  font-size: 12px;
  color: #707481;
}
.dlg-btn {
  border: 1px solid #d8dbde;
  background: #fff;
  color: #1f2329;
  padding: 6px 14px;
  border-radius: 5px;
  cursor: pointer;
  font-size: 13px;
}
.dlg-btn:hover { background: #f5f7fa; }
.dlg-btn.primary { background: #21a366; border-color: #21a366; color: #fff; }
.dlg-btn.primary:hover { background: #1a8c56; }
.dlg-btn.ghost { background: transparent; }
</style>
