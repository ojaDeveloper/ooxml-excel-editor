<script setup lang="ts">
import { computed, defineComponent, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, watch } from 'vue'
import type * as EChartsNS from 'echarts'
import type { ExcelPlugin, ExcelPluginContext, OverlayContext, PluginEvent, ToolbarItem, ViewerApi } from '@/core/plugin'
import type { ExcelSource } from '@/core/loader'
import type {
  CellModel,
  CellStyleFn,
  ChartSpec,
  ImageAnchor,
  MergeRange,
  SheetModel,
  TransformModelFn,
  WorkbookModel,
} from '@/core/model/types'
import type { ParseProgress } from '@/core/progress'
import type { ReadOptions } from '@/core/model/data-access'
import { getCellValue, getCellText, getSheetData, getRangeData, sheetToJSON } from '@/core/model/data-access'
import type { ViewerTheme } from '@/core/render/theme'
import { useExcelDocument } from '@/composables/useExcelDocument'
import { CanvasRenderer, type ViewState } from '@/core/render/canvas-renderer'
import { GridMetrics, colIndexToLetters } from '@/core/layout/grid-metrics'
import { anchorRect } from '@/core/overlay/anchor'
import { chartToOption } from '@/core/overlay/chart-mapper'
import { loadECharts } from '@/core/overlay/echarts-loader'
import { ViewerController } from '@/core/viewer/controller'
import { revokeImages } from '@/core/finalize'
import {
  canvasToBlob,
  compositeOverlays,
  downloadBlob,
  exportToPdf,
  exportToVectorPdf,
  loadImage,
  printSheets,
  type ExportDecorations,
  type ExportSheetImage,
  type ExportTarget,
  type ImageExportOptions,
  type PdfExportOptions,
  type PrintOptions,
  type VectorSheet,
} from '@/core/export'
import ViewerToolbar from './ViewerToolbar.vue'
import SheetTabs from './SheetTabs.vue'
import ExportDialog from './ExportDialog.vue'
import FindBar from './FindBar.vue'
import FilterPopup from './FilterPopup.vue'
import ActionToolbar from './ActionToolbar.vue'
import type { ExportConfig } from './export-types'
import type { ResolvedToolbarItem } from './toolbar-types'
import { TOOLBAR_ICONS } from './toolbar-icons'

const props = withDefaults(
  defineProps<{
    src?: ExcelSource
    fileName?: string
    /** 外观主题(覆盖默认配色) */
    theme?: Partial<ViewerTheme>
    /** 数据钩子: 解析后改模型再渲染 */
    transformModel?: TransformModelFn
    /** 渲染钩子: 按单元格覆盖样式 */
    cellStyle?: CellStyleFn
    /** 单击超链接是否默认在新标签打开(false 时只派发 hyperlink-click 事件) */
    openLinks?: boolean
    /** 插件列表(打包主题/钩子/事件/overlay) */
    plugins?: ExcelPlugin[]
    /**
     * 操作工具栏(顶栏下一行): 内置 'find'/'filter'/'sort' 默认显示。
     * false 隐藏整条;数组显式控制项与顺序(内置 id 或自定义 ToolbarItem)。
     * 插件 ExcelPlugin.toolbar 贡献的项总会追加(opt-in)。
     */
    toolbar?: boolean | Array<string | ToolbarItem>
  }>(),
  // toolbar 默认 true(显示内置项);若不显式给默认,Vue 会把布尔型 prop 缺省判成 false
  { openLinks: true, toolbar: true },
)

const normalizedPlugins = computed<ExcelPlugin[]>(() => props.plugins ?? [])

// 合并各扩展点: 插件按数组顺序,组件 prop 最后覆盖
const effectiveTheme = computed(() =>
  Object.assign({}, ...normalizedPlugins.value.map((p) => p.theme || {}), props.theme || {}),
)
const hasCellStyleHook = computed(() => !!props.cellStyle || normalizedPlugins.value.some((p) => p.cellStyle))
function combinedCellStyle(cell: CellModel, pos: { row: number; col: number }) {
  let acc: Record<string, unknown> | undefined
  const apply = (fn?: CellStyleFn) => {
    const o = fn?.(cell, pos)
    if (o) acc = { ...(acc || {}), ...o }
  }
  for (const p of normalizedPlugins.value) apply(p.cellStyle)
  apply(props.cellStyle)
  return acc as ReturnType<CellStyleFn>
}
function effectiveTransform(wb: WorkbookModel): WorkbookModel {
  let m = wb
  for (const p of normalizedPlugins.value) if (p.transformModel) m = p.transformModel(m) ?? m
  if (props.transformModel) m = props.transformModel(m) ?? m
  return m
}

const emit = defineEmits<{
  /** 工作簿解析并首次渲染完成 */
  (e: 'rendered', workbook: WorkbookModel): void
  /** 解析失败(友好错误文案) */
  (e: 'error', message: string): void
  /** 解析进度(分阶段) */
  (e: 'progress', progress: ParseProgress): void
  /** 单击单元格 */
  (e: 'cell-click', payload: { row: number; col: number; text: string }): void
  /** 双击单元格 */
  (e: 'cell-dblclick', payload: { row: number; col: number; text: string }): void
  /** 选区变化 */
  (e: 'selection-change', payload: { range: MergeRange; active: { row: number; col: number } }): void
  /** 切换工作表 */
  (e: 'sheet-change', payload: { index: number; name: string }): void
  /** 单击超链接(openLinks=false 时由你处理跳转) */
  (e: 'hyperlink-click', payload: { url: string; cell: { row: number; col: number } }): void
}>()

const { loading, error, workbook, load, progress } = useExcelDocument()

const progressLabel = computed(() => {
  const p = progress.value
  if (!p) return ''
  return p.stage === 'read' ? '读取文件…' : p.stage === 'parse' ? '解析中…' : '构建表格…'
})
const progressPct = computed(() => {
  const p = progress.value
  return p && p.ratio != null ? Math.round(p.ratio * 100) : null
})

const activeSheet = ref(0)
const zoom = ref(1)

const rootEl = ref<HTMLElement | null>(null)
const renderAreaEl = ref<HTMLElement | null>(null)
const canvasEl = ref<HTMLCanvasElement | null>(null)
const scrollerEl = ref<HTMLElement | null>(null)
// 叠加层四象限容器(冻结窗格): 主区 / 冻结行 / 冻结列 / 冻结角
const ovMain = ref<HTMLElement | null>(null)
const ovFRow = ref<HTMLElement | null>(null)
const ovFCol = ref<HTMLElement | null>(null)
const ovCorner = ref<HTMLElement | null>(null)

const renderer = shallowRef<CanvasRenderer | null>(null)
const view = ref<ViewState>({ scrollX: 0, scrollY: 0, width: 0, height: 0, zoom: 1 })

const sheet = computed<SheetModel | null>(() => {
  const wb = workbook.value
  if (!wb) return null
  return wb.sheets[activeSheet.value] ?? wb.sheets[0] ?? null
})

// ---------------- 渲染引擎: 委托框架无关 ViewerController ----------------
// overlay slot 用: 每次重绘 +1 → 作用域插槽重算 rectOf 位置(随滚动/缩放/切表跟随)
const renderTick = ref(0)
const spacerEl = ref<HTMLElement | null>(null)
let controller: ViewerController | null = null

function doRender() {
  controller?.render()
}
function scheduleRender() {
  controller?.scheduleRender()
}
function measure() {
  controller?.measure()
}

function rebuildRenderer() {
  const s = sheet.value
  const wb = workbook.value
  if (!s || !wb || !controller) return
  // 先清 Vue 侧状态,再重建(rebuild 末尾会以"已清空选区"绘制)
  clearSelection()
  findHits.value = []
  findIndex.value = -1
  tooltip.value = null
  controller.rebuild(s, wb, zoom.value, {
    theme: effectiveTheme.value,
    cellStyle: hasCellStyleHook.value ? combinedCellStyle : undefined,
  })
  if (findQuery.value) recomputeFind() // 新表上重新查找
}

function onScroll() {
  const sc = scrollerEl.value
  if (!sc) return
  tooltip.value = null
  controller?.setScroll(sc.scrollLeft, sc.scrollTop)
}

let resizeObserver: ResizeObserver | null = null

onMounted(() => {
  initPlugins()
  // DOM 挂载后实例化框架无关控制器,把 canvas/scroller/spacer/叠加层四象限交给它
  if (canvasEl.value && renderAreaEl.value && scrollerEl.value && spacerEl.value && ovMain.value && ovFRow.value && ovFCol.value && ovCorner.value) {
    controller = new ViewerController(
      {
        canvas: canvasEl.value,
        renderArea: renderAreaEl.value,
        scroller: scrollerEl.value,
        spacer: spacerEl.value,
        overlays: { main: ovMain.value, frow: ovFRow.value, fcol: ovFCol.value, corner: ovCorner.value },
      },
      {
        getSelection: () => selection.value,
        onRenderer: (r) => (renderer.value = r),
        onRenderTick: () => renderTick.value++,
      },
    )
    view.value = controller.view // 壳与控制器共享同一 view 对象(现有 view.value 读法不变)
  }
  if (props.src) load(props.src, effectiveTransform)
  resizeObserver = new ResizeObserver(() => {
    measure()
    doRender()
  })
  if (renderAreaEl.value) resizeObserver.observe(renderAreaEl.value)
})

onBeforeUnmount(() => {
  resizeObserver?.disconnect()
  controller?.dispose()
  pluginCleanups.forEach((fn) => fn())
  if (workbook.value) revokeImages(workbook.value)
})

watch(() => props.src, (s) => {
  if (s) load(s, effectiveTransform)
})

// 主题 / cellStyle / 插件 变化 → 重建渲染器(它们在构造时注入)
watch(
  () => [effectiveTheme.value, props.cellStyle, props.plugins],
  () => {
    if (renderer.value) rebuildRenderer()
  },
  { deep: true },
)

// 插件列表变化 → 重新初始化生命周期/事件订阅
watch(() => props.plugins, () => initPlugins(), { deep: false })

watch(workbook, async (wb) => {
  if (!wb) return
  // 新工作簿: 旧筛选态作废(模型已换)
  filterOrigHidden.clear()
  filterState.clear()
  filterPopup.value = null
  activeSheet.value = wb.activeSheet
  await nextTick()
  rebuildRenderer()
  emit('rendered', wb)
})

// 切表派发 sheet-change
watch(activeSheet, (idx) => {
  const wb = workbook.value
  if (wb?.sheets[idx]) fire('sheet-change', { index: idx, name: wb.sheets[idx].name })
})

watch(error, (msg) => {
  if (msg) emit('error', msg)
})

watch(progress, (p) => {
  if (p) emit('progress', p)
})

watch(activeSheet, async (_idx, oldIdx) => {
  resetFilterFor(oldIdx) // 离开旧表: 恢复其筛选隐藏的行
  await nextTick()
  rebuildRenderer()
})

watch(zoom, (z) => controller?.setZoom(z))

// ---------------- 交互: 选区 / 超链接 / 悬停 / 复制 ----------------
type Cell = { row: number; col: number }
const selAnchor = ref<Cell | null>(null) // 固定角(扩选时不动)
const selActive = ref<Cell | null>(null) // 活动角(移动/扩选时变)
const selMode = ref<'range' | 'rows' | 'cols'>('range')
const tooltip = ref<{ text: string; x: number; y: number; kind: 'overflow' | 'comment' } | null>(null)
let dragMode: 'none' | 'cell' | 'row' | 'col' | 'resize-col' | 'resize-row' = 'none'
let resizeTarget = -1 // 正在拖拽改宽高的列/行索引
let resizeStartPos = 0 // 起始鼠标坐标(px)
let resizeStartSize = 0 // 起始宽/高(px)
let dragMoved = false

function cellRange(c: Cell): MergeRange {
  return { top: c.row, left: c.col, bottom: c.row, right: c.col }
}

const selection = computed<MergeRange | null>(() => {
  const r = renderer.value
  const a = selAnchor.value
  const b = selActive.value
  if (!r || !a || !b) return null
  if (selMode.value === 'rows') {
    return { top: Math.min(a.row, b.row), bottom: Math.max(a.row, b.row), left: 0, right: r.metrics.cols - 1 }
  }
  if (selMode.value === 'cols') {
    return { left: Math.min(a.col, b.col), right: Math.max(a.col, b.col), top: 0, bottom: r.metrics.rows - 1 }
  }
  const ra = r.mergeAt(a.row, a.col) ?? cellRange(a)
  const rb = r.mergeAt(b.row, b.col) ?? cellRange(b)
  return {
    top: Math.min(ra.top, rb.top),
    left: Math.min(ra.left, rb.left),
    bottom: Math.max(ra.bottom, rb.bottom),
    right: Math.max(ra.right, rb.right),
  }
})

const activeCellAddr = computed(() => {
  const c = selActive.value
  return c ? colIndexToLetters(c.col) + (c.row + 1) : ''
})
const selRangeLabel = computed(() => {
  const s = selection.value
  if (!s || (s.top === s.bottom && s.left === s.right)) return ''
  return `${colIndexToLetters(s.left)}${s.top + 1}:${colIndexToLetters(s.right)}${s.bottom + 1}`
})
const formulaBarText = computed(() => {
  const r = renderer.value
  const c = selActive.value
  if (!r || !c) return ''
  return r.cellFormula(c.row, c.col) ?? r.cellText(c.row, c.col)
})
const stats = computed(() => {
  const r = renderer.value
  const s = selection.value
  return r && s ? r.selectionStats(s) : null
})

function fmtNum(n: number): string {
  if (!isFinite(n)) return '—'
  return n.toLocaleString('en-US', { maximumFractionDigits: 2 })
}

function clearSelection() {
  selAnchor.value = null
  selActive.value = null
  selMode.value = 'range'
}

// ---- 命中区域 ----
type Hit =
  | { region: 'cell'; row: number; col: number }
  | { region: 'row'; row: number }
  | { region: 'col'; col: number }
  | { region: 'corner' }
  | { region: 'none' }

function localXY(e: MouseEvent): { x: number; y: number } | null {
  const area = renderAreaEl.value
  if (!area) return null
  const rect = area.getBoundingClientRect()
  return { x: e.clientX - rect.left, y: e.clientY - rect.top }
}
function hitRegion(e: MouseEvent): Hit {
  const r = renderer.value
  const p = localXY(e)
  if (!r || !p) return { region: 'none' }
  const hw = r.metrics.rowHeaderWidth
  const hh = r.metrics.colHeaderHeight
  if (p.x < hw && p.y < hh) return { region: 'corner' }
  if (p.x < hw) return { region: 'row', row: r.rowAtScreen(view.value, p.y) }
  if (p.y < hh) return { region: 'col', col: r.colAtScreen(view.value, p.x) }
  const cell = r.cellAtScreen(view.value, p.x, p.y)
  return cell ? { region: 'cell', row: cell.row, col: cell.col } : { region: 'none' }
}

// ---- 选区设置 ----
function setCell(cell: Cell, extend: boolean) {
  selMode.value = 'range'
  if (extend && selAnchor.value) selActive.value = cell
  else {
    selAnchor.value = cell
    selActive.value = cell
  }
}
function setRows(row: number, extend: boolean) {
  selMode.value = 'rows'
  const c: Cell = { row, col: 0 }
  if (extend && selAnchor.value) selActive.value = c
  else {
    selAnchor.value = c
    selActive.value = c
  }
}
function setCols(col: number, extend: boolean) {
  selMode.value = 'cols'
  const c: Cell = { row: 0, col }
  if (extend && selAnchor.value) selActive.value = c
  else {
    selAnchor.value = c
    selActive.value = c
  }
}
function selectAll() {
  const r = renderer.value
  if (!r) return
  selMode.value = 'range'
  selAnchor.value = { row: 0, col: 0 }
  selActive.value = { row: r.metrics.rows - 1, col: r.metrics.cols - 1 }
}

// ---- 鼠标 ----
function onMouseDown(e: MouseEvent) {
  if (e.button !== 0) return
  scrollerEl.value?.focus()
  const r = renderer.value
  const p = localXY(e)
  // 自动筛选下拉按钮(优先于一切)
  if (r && p) {
    const fcol = r.filterButtonAt(view.value, p.x, p.y)
    if (fcol != null) {
      openFilterPopup(fcol)
      return
    }
  }
  // 表头边界拖拽改宽高(优先于选择)
  if (r && p) {
    if (p.y < r.metrics.colHeaderHeight) {
      const b = nearColBorder(p.x, p.y)
      if (b) {
        dragMode = 'resize-col'
        resizeTarget = b.col
        resizeStartPos = p.x
        resizeStartSize = r.metrics.colWidth(b.col)
        return
      }
    } else if (p.x < r.metrics.rowHeaderWidth) {
      const b = nearRowBorder(p.x, p.y)
      if (b) {
        dragMode = 'resize-row'
        resizeTarget = b.row
        resizeStartPos = p.y
        resizeStartSize = r.metrics.rowHeight(b.row)
        return
      }
    }
  }
  const hit = hitRegion(e)
  dragMoved = false
  if (hit.region === 'corner') {
    selectAll()
    dragMode = 'none'
  } else if (hit.region === 'row') {
    dragMode = 'row'
    setRows(hit.row, e.shiftKey)
  } else if (hit.region === 'col') {
    dragMode = 'col'
    setCols(hit.col, e.shiftKey)
  } else if (hit.region === 'cell') {
    dragMode = 'cell'
    setCell({ row: hit.row, col: hit.col }, e.shiftKey)
  } else {
    dragMode = 'none'
  }
  doRender()
}
function onMouseMove(e: MouseEvent) {
  if (dragMode !== 'none') {
    const r = renderer.value
    const p = localXY(e)
    if (!r || !p) return
    dragMoved = true
    if (dragMode === 'resize-col') {
      controller?.setColWidthPx(resizeTarget, resizeStartSize + (p.x - resizeStartPos))
      return
    }
    if (dragMode === 'resize-row') {
      controller?.setRowHeightPx(resizeTarget, resizeStartSize + (p.y - resizeStartPos))
      return
    }
    if (dragMode === 'cell') {
      const cell = r.cellAtScreen(view.value, p.x, p.y)
      if (cell) {
        selActive.value = cell
        scheduleRender()
      }
    } else if (dragMode === 'row') {
      const row = r.rowAtScreen(view.value, p.y)
      if (row >= 0) {
        selActive.value = { row, col: 0 }
        scheduleRender()
      }
    } else {
      const col = r.colAtScreen(view.value, p.x)
      if (col >= 0) {
        selActive.value = { row: 0, col }
        scheduleRender()
      }
    }
    return
  }
  updateHover(e)
}
function onMouseUp(e: MouseEvent) {
  if (dragMode === 'cell' && !dragMoved) {
    const hit = hitRegion(e)
    const r = renderer.value
    if (hit.region === 'cell' && r) {
      fire('cell-click', { row: hit.row, col: hit.col, text: r.cellText(hit.row, hit.col) })
      const link = r.cellHyperlink(hit.row, hit.col)
      if (link) {
        fire('hyperlink-click', { url: link, cell: { row: hit.row, col: hit.col } })
        if (props.openLinks) window.open(link, '_blank', 'noopener')
      }
    }
  }
  dragMode = 'none'
}
function updateHover(e: MouseEvent) {
  const r = renderer.value
  const sc = scrollerEl.value
  const p = localXY(e)
  if (!r || !sc || !p) {
    tooltip.value = null
    return
  }
  // 表头边界 → 改宽高光标
  if (p.y < r.metrics.colHeaderHeight && nearColBorder(p.x, p.y)) {
    sc.style.cursor = 'col-resize'
    tooltip.value = null
    return
  }
  if (p.x < r.metrics.rowHeaderWidth && nearRowBorder(p.x, p.y)) {
    sc.style.cursor = 'row-resize'
    tooltip.value = null
    return
  }
  const cell = r.cellAtScreen(view.value, p.x, p.y)
  if (!cell) {
    tooltip.value = null
    sc.style.cursor = ''
    return
  }
  sc.style.cursor = r.cellHyperlink(cell.row, cell.col) ? 'pointer' : 'cell'
  const tx = p.x + 14
  const ty = p.y + 18
  const comment = r.commentAt(cell.row, cell.col)
  if (comment) {
    tooltip.value = { text: comment, x: tx, y: ty, kind: 'comment' }
    return
  }
  const full = r.overflowTextAt(cell.row, cell.col)
  tooltip.value = full ? { text: full, x: tx, y: ty, kind: 'overflow' } : null
}
function onMouseLeave() {
  tooltip.value = null
}

// ---- 宽高自适应(双击边界) / 边界命中 ----
function nearColBorder(x: number, y: number): { col: number } | null {
  const r = renderer.value
  if (!r || y >= r.metrics.colHeaderHeight) return null
  const col = r.colAtScreen(view.value, x)
  if (col < 0) return null
  const rect = r.screenRectOfCell(view.value, 0, col)
  if (Math.abs(x - (rect.x + rect.w)) <= 4) return { col }
  if (Math.abs(x - rect.x) <= 4 && col > 0) return { col: col - 1 }
  return null
}
function nearRowBorder(x: number, y: number): { row: number } | null {
  const r = renderer.value
  if (!r || x >= r.metrics.rowHeaderWidth) return null
  const row = r.rowAtScreen(view.value, y)
  if (row < 0) return null
  const rect = r.screenRectOfCell(view.value, row, 0)
  if (Math.abs(y - (rect.y + rect.h)) <= 4) return { row }
  if (Math.abs(y - rect.y) <= 4 && row > 0) return { row: row - 1 }
  return null
}
function onDblClick(e: MouseEvent) {
  const r = renderer.value
  const p = localXY(e)
  if (!r || !p) return
  const colHit = nearColBorder(p.x, p.y)
  const rowHit = colHit ? null : nearRowBorder(p.x, p.y)
  if (colHit) {
    controller?.autoFitColumn(colHit.col)
  } else if (rowHit) {
    controller?.autoFitRow(rowHit.row)
  } else {
    // 非边界 → 双击单元格事件
    const cell = r.cellAtScreen(view.value, p.x, p.y)
    if (cell) fire('cell-dblclick', { row: cell.row, col: cell.col, text: r.cellText(cell.row, cell.col) })
  }
}

// ---- 键盘 ----
function pageRows(): number {
  const r = renderer.value
  if (!r) return 10
  return Math.max(1, Math.floor((view.value.height - r.metrics.colHeaderHeight) / r.defaultRowPx) - 1)
}
function onKeyDown(e: KeyboardEvent) {
  if ((e.ctrlKey || e.metaKey) && (e.key === 'c' || e.key === 'C')) {
    copySelection()
    e.preventDefault()
    return
  }
  if ((e.ctrlKey || e.metaKey) && (e.key === 'a' || e.key === 'A')) {
    selectAll()
    doRender()
    e.preventDefault()
    return
  }
  const r = renderer.value
  if (!r || !selActive.value) return
  const maxRow = r.metrics.rows - 1
  const maxCol = r.metrics.cols - 1
  const ctrl = e.ctrlKey || e.metaKey
  let { row, col } = selActive.value
  let handled = true
  switch (e.key) {
    case 'ArrowUp': if (ctrl) { const j = jumpEdge(r, row, col, -1, 0); row = j.row; col = j.col } else row = Math.max(0, row - 1); break
    case 'ArrowDown': if (ctrl) { const j = jumpEdge(r, row, col, 1, 0); row = j.row; col = j.col } else row = Math.min(maxRow, row + 1); break
    case 'ArrowLeft': if (ctrl) { const j = jumpEdge(r, row, col, 0, -1); row = j.row; col = j.col } else col = Math.max(0, col - 1); break
    case 'ArrowRight': if (ctrl) { const j = jumpEdge(r, row, col, 0, 1); row = j.row; col = j.col } else col = Math.min(maxCol, col + 1); break
    case 'Home': col = 0; if (ctrl) row = 0; break
    case 'End': col = maxCol; if (e.ctrlKey) row = maxRow; break
    case 'PageUp': row = Math.max(0, row - pageRows()); break
    case 'PageDown': row = Math.min(maxRow, row + pageRows()); break
    case 'Enter': row = Math.min(maxRow, row + 1); break
    case 'Tab': col = e.shiftKey ? Math.max(0, col - 1) : Math.min(maxCol, col + 1); break
    default: handled = false
  }
  if (!handled) return
  e.preventDefault()
  const m = r.mergeAt(row, col)
  if (m) {
    row = m.top
    col = m.left
  }
  selMode.value = 'range'
  selActive.value = { row, col }
  const extend = e.shiftKey && e.key !== 'Tab'
  if (!extend) selAnchor.value = { row, col }
  scrollActiveIntoView()
  doRender()
}
function scrollActiveIntoView() {
  const r = renderer.value
  const sc = scrollerEl.value
  const c = selActive.value
  if (!r || !sc || !c) return
  const hw = r.metrics.rowHeaderWidth
  const hh = r.metrics.colHeaderHeight
  const fz = r.freezeGeometry
  let sx = sc.scrollLeft
  let sy = sc.scrollTop
  if (c.col >= fz.frozenCols) {
    const cl = r.metrics.colLeft(c.col)
    const cr = cl + r.metrics.colWidth(c.col)
    const viewW = view.value.width - hw
    if (cr > sx + viewW) sx = cr - viewW
    if (cl < sx + fz.frozenWidth) sx = cl - fz.frozenWidth
  }
  if (c.row >= fz.frozenRows) {
    const ct = r.metrics.rowTop(c.row)
    const cb = ct + r.metrics.rowHeight(c.row)
    const viewH = view.value.height - hh
    if (cb > sy + viewH) sy = cb - viewH
    if (ct < sy + fz.frozenHeight) sy = ct - fz.frozenHeight
  }
  sx = Math.max(0, sx)
  sy = Math.max(0, sy)
  if (sx !== sc.scrollLeft || sy !== sc.scrollTop) {
    sc.scrollLeft = sx
    sc.scrollTop = sy
    view.value.scrollX = sx
    view.value.scrollY = sy
  }
}
/** Ctrl+方向: 跳到数据块边界(Excel 行为) */
function jumpEdge(
  r: CanvasRenderer,
  row: number,
  col: number,
  dr: number,
  dc: number,
): { row: number; col: number } {
  const maxRow = r.metrics.rows - 1
  const maxCol = r.metrics.cols - 1
  const inB = (rr: number, cc: number) => rr >= 0 && rr <= maxRow && cc >= 0 && cc <= maxCol
  const filled = (rr: number, cc: number) => r.cellText(rr, cc) !== ''
  let nr = row + dr
  let nc = col + dc
  if (!inB(nr, nc)) return { row, col }
  if (filled(row, col) && filled(nr, nc)) {
    // 沿填充块走到块尾
    while (inB(nr + dr, nc + dc) && filled(nr + dr, nc + dc)) {
      nr += dr
      nc += dc
    }
  } else {
    // 跳过空白到下一个填充(或边界)
    while (inB(nr, nc) && !filled(nr, nc)) {
      if (!inB(nr + dr, nc + dc)) break
      nr += dr
      nc += dc
    }
  }
  return { row: nr, col: nc }
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}
async function copySelection() {
  const r = renderer.value
  const s = selection.value
  if (!r || !s) return
  // 防超大选区卡死: 复制范围软上限
  const rowEnd = Math.min(s.bottom, s.top + 4999)
  const colEnd = Math.min(s.right, s.left + 255)
  const lines: string[] = []
  const htmlRows: string[] = []
  for (let row = s.top; row <= rowEnd; row++) {
    const cells: string[] = []
    const htmlCells: string[] = []
    for (let col = s.left; col <= colEnd; col++) {
      const text = r.cellText(row, col)
      cells.push(text)
      const css = r.cellInlineStyle(row, col)
      htmlCells.push(`<td${css ? ` style="${css}"` : ''}>${escapeHtml(text)}</td>`)
    }
    lines.push(cells.join('\t'))
    htmlRows.push(`<tr>${htmlCells.join('')}</tr>`)
  }
  const tsv = lines.join('\n')
  const html = `<table border="1" style="border-collapse:collapse">${htmlRows.join('')}</table>`
  try {
    // 优先写 text/plain + text/html(粘到 Word/Excel 保留表格与格式)
    const ClipItem = (window as any).ClipboardItem
    if (ClipItem && navigator.clipboard?.write) {
      await navigator.clipboard.write([
        new ClipItem({
          'text/plain': new Blob([tsv], { type: 'text/plain' }),
          'text/html': new Blob([html], { type: 'text/html' }),
        }),
      ])
    } else {
      await navigator.clipboard.writeText(tsv)
    }
  } catch {
    /* 某些环境无剪贴板权限，静默忽略 */
  }
}

// ---------------- 选区变化事件 + overlay 定位 API ----------------
// selection/selActive 在上方已定义,这里安全引用
watch(selection, (sel) => {
  if (sel && selActive.value) fire('selection-change', { range: sel, active: selActive.value })
})

/** 单元格当前屏幕矩形(render-area 相对坐标);供 overlay slot / 命令式定位 */
function rectOf(row: number, col: number): { x: number; y: number; w: number; h: number } | null {
  return controller?.rectOf(row, col) ?? null
}
/** 区域当前屏幕矩形(左上到右下的并集) */
function rectOfRange(range: MergeRange): { x: number; y: number; w: number; h: number } | null {
  return controller?.rectOfRange(range) ?? null
}

// ---------------- 导出 / 打印 ----------------
/** target → 工作表索引列表 */
function resolveTargets(target: ExportTarget = 'active'): number[] {
  const wb = workbook.value
  if (!wb) return []
  if (target === 'all') return wb.sheets.map((_, i) => i).filter((i) => wb.sheets[i].state === 'visible')
  if (target === 'active') return [activeSheet.value]
  if (typeof target === 'number') return [target]
  return target.filter((i) => wb.sheets[i])
}

/** 离屏渲染一个图表为 dataURL(供非当前表 / 统一合成);echarts 不可用返回 null */
async function chartDataUrl(spec: ChartSpec, metrics: GridMetrics): Promise<string | null> {
  let echarts: typeof EChartsNS
  try {
    echarts = await loadECharts()
  } catch {
    return null
  }
  const rect = anchorRect(metrics, spec.anchor)
  const div = document.createElement('div')
  div.style.cssText = `position:fixed;left:-10000px;top:0;width:${Math.max(80, Math.round(rect.width))}px;height:${Math.max(60, Math.round(rect.height))}px`
  document.body.appendChild(div)
  const inst = echarts.init(div)
  try {
    inst.setOption(chartToOption(spec))
    return inst.getDataURL({ pixelRatio: 2, backgroundColor: '#fff' })
  } catch {
    return null
  } finally {
    inst.dispose()
    div.remove()
  }
}

/** 收集一个工作表的叠加层装饰(图片/图表/形状),供合成到导出底图 */
async function collectDecorations(s: SheetModel, metrics: GridMetrics): Promise<ExportDecorations> {
  const images: { source: CanvasImageSource; anchor: ImageAnchor }[] = []
  for (const anchor of s.images) {
    if (!anchor.src) continue
    try {
      images.push({ source: await loadImage(anchor.src), anchor })
    } catch {
      /* 单张图加载失败跳过 */
    }
  }
  const charts: { source: CanvasImageSource; anchor: ImageAnchor }[] = []
  for (const chart of s.charts) {
    const url = await chartDataUrl(chart, metrics)
    if (!url) continue
    try {
      charts.push({ source: await loadImage(url), anchor: chart.anchor })
    } catch {
      /* 跳过 */
    }
  }
  return { images, charts, shapes: s.shapes }
}

/**
 * 为一个工作表生成合成底图(格子 + 图片/图表/形状)。
 * withTitles(PDF/打印): 应用 pageSetup 打印标题行(抽出标题条 + 正文剔除标题行)与缩放。
 * 范围优先级: opts.range > pageSetup.printArea > 整表。
 */
async function buildSheetImage(
  sheetIdx: number,
  opts: PdfExportOptions,
  withTitles = false,
): Promise<ExportSheetImage | null> {
  const wb = workbook.value
  const s = wb?.sheets[sheetIdx]
  if (!wb || !s) return null
  // 当前表复用 live renderer;其它表临时建一个(在离屏 canvas 上)
  const r =
    sheetIdx === activeSheet.value && renderer.value
      ? renderer.value
      : new CanvasRenderer(document.createElement('canvas'), s, wb, 1, {
          theme: effectiveTheme.value,
          cellStyle: hasCellStyleHook.value ? combinedCellStyle : undefined,
        })

  const ps = s.pageSetup
  const dim = s.dimension
  const full: MergeRange = { top: 0, left: 0, bottom: Math.max(0, dim.rows - 1), right: Math.max(0, dim.cols - 1) }
  const range = opts.range ?? ps?.printArea ?? full

  // 打印标题行/列: 标题在正文上方/左侧时,抽出标题条,正文起点相应内移(避免重复)
  let titleRows: [number, number] | null = null
  let titleCols: [number, number] | null = null
  let bodyTop = range.top
  let bodyLeft = range.left
  if (withTitles && ps?.printTitleRows) {
    const [a, b] = ps.printTitleRows
    if (b >= a && a <= range.top && b < range.bottom) {
      titleRows = [a, b]
      bodyTop = Math.max(range.top, b + 1)
    }
  }
  if (withTitles && ps?.printTitleCols) {
    const [a, b] = ps.printTitleCols
    if (b >= a && a <= range.left && b < range.right) {
      titleCols = [a, b]
      bodyLeft = Math.max(range.left, b + 1)
    }
  }

  const render = (rg: MergeRange, scale?: number) =>
    r.exportToCanvas({
      range: rg,
      scale: scale ?? opts.scale,
      includeHeaders: opts.includeHeaders,
      gridlines: opts.gridlines,
      background: opts.background,
    })

  const base = render({ top: bodyTop, left: bodyLeft, bottom: range.bottom, right: range.right })
  const deco = await collectDecorations(s, base.metrics)
  compositeOverlays(base, deco)
  const S = base.scale // 标题条用同一 scale → 与正文等宽/等高,可逐页拼接

  let repeatTop: ExportSheetImage['repeatTop']
  let repeatLeft: ExportSheetImage['repeatLeft']
  let corner: ExportSheetImage['corner']
  if (titleRows) {
    const strip = render({ top: titleRows[0], bottom: titleRows[1], left: bodyLeft, right: range.right }, S)
    repeatTop = { canvas: strip.canvas, heightCss: strip.bodyH }
  }
  if (titleCols) {
    const strip = render({ top: bodyTop, bottom: range.bottom, left: titleCols[0], right: titleCols[1] }, S)
    repeatLeft = { canvas: strip.canvas, widthCss: strip.bodyW }
  }
  if (titleRows && titleCols) {
    const c = render({ top: titleRows[0], bottom: titleRows[1], left: titleCols[0], right: titleCols[1] }, S)
    corner = { canvas: c.canvas }
  }

  // 打印缩放: 非 fitToWidth 时套用 pageSetup.scale
  const zoom = opts.fitToWidth === false && ps?.scale ? ps.scale / 100 : undefined

  return {
    canvas: base.canvas,
    bodyWcss: base.bodyW,
    bodyHcss: base.bodyH,
    sheetName: s.name,
    repeatTop,
    repeatLeft,
    corner,
    zoom,
  }
}

/** 为一个工作表生成矢量导出输入(逐格信息 + 兜底底图 + 图片图表 + 标题行) */
async function buildVectorSheet(sheetIdx: number, opts: PdfExportOptions): Promise<VectorSheet | null> {
  const wb = workbook.value
  const s = wb?.sheets[sheetIdx]
  if (!wb || !s) return null
  const r =
    sheetIdx === activeSheet.value && renderer.value
      ? renderer.value
      : new CanvasRenderer(document.createElement('canvas'), s, wb, 1, {
          theme: effectiveTheme.value,
          cellStyle: hasCellStyleHook.value ? combinedCellStyle : undefined,
        })
  const metrics = new GridMetrics(s, 1)
  const ps = s.pageSetup
  const dim = s.dimension
  const full: MergeRange = { top: 0, left: 0, bottom: Math.max(0, dim.rows - 1), right: Math.max(0, dim.cols - 1) }
  const range = opts.range ?? ps?.printArea ?? full

  let titleRows: [number, number] | undefined
  let titleCols: [number, number] | undefined
  let bodyTop = range.top
  let bodyLeft = range.left
  if (ps?.printTitleRows) {
    const [a, b] = ps.printTitleRows
    if (b >= a && a <= range.top && b < range.bottom) {
      titleRows = [a, b]
      bodyTop = Math.max(range.top, b + 1)
    }
  }
  if (ps?.printTitleCols) {
    const [a, b] = ps.printTitleCols
    if (b >= a && a <= range.left && b < range.right) {
      titleCols = [a, b]
      bodyLeft = Math.max(range.left, b + 1)
    }
  }
  // 兜底底图须覆盖 标题行/列 ∪ 正文(标题格无字体中文也要能裁图)
  const rasterTop = titleRows ? Math.min(range.top, titleRows[0]) : range.top
  const rasterLeft = titleCols ? Math.min(range.left, titleCols[0]) : range.left

  const base = r.exportToCanvas({
    range: { top: rasterTop, left: rasterLeft, bottom: range.bottom, right: range.right },
    scale: opts.scale,
    includeHeaders: false,
    gridlines: opts.gridlines,
    background: opts.background,
  })
  const deco = await collectDecorations(s, base.metrics)
  const images = [...(deco.images ?? []), ...(deco.charts ?? [])]
  const zoom = opts.fitToWidth === false && ps?.scale ? ps.scale / 100 : undefined

  return {
    sheetName: s.name,
    metrics,
    bodyLeft,
    bodyRight: range.right,
    bodyTop,
    bodyBottom: range.bottom,
    titleRows,
    titleCols,
    merges: s.merges,
    gridlines: opts.gridlines ?? s.showGridLines,
    zoom,
    getCell: (rr, cc) => r.exportCellDraw(rr, cc),
    rasterCanvas: base.canvas,
    rasterScale: base.scale,
    rasterTop,
    rasterLeft,
    images,
  }
}

/** 从工作表原生 pageSetup 推导导出默认值(纸张/方向/边距/是否适应页宽) */
function pageSetupDefaults(sheetIdx: number): Partial<PdfExportOptions> {
  const ps = workbook.value?.sheets[sheetIdx]?.pageSetup
  if (!ps) return {}
  const d: Partial<PdfExportOptions> = {}
  if (ps.paperFormat) d.format = ps.paperFormat
  if (ps.orientation) d.orientation = ps.orientation
  if (ps.margins) d.margin = { top: ps.margins.top, right: ps.margins.right, bottom: ps.margins.bottom, left: ps.margins.left }
  // fitToPage → 适应页宽;否则按自然尺寸×scale(下面 buildSheetImage 用 zoom 处理)
  d.fitToWidth = ps.fitToPage ? true : false
  return d
}

function baseName(): string {
  return (props.fileName || workbook.value?.sheets[activeSheet.value]?.name || 'workbook').replace(/\.[^.]+$/, '')
}

/** 导出当前/指定表为图片 Blob(图片为单表;多表请用 PDF) */
async function exportImage(opts: ImageExportOptions = {}): Promise<Blob> {
  const targets = resolveTargets(opts.target)
  if (!targets.length) throw new Error('无可导出的工作表')
  const img = await buildSheetImage(targets[0], opts)
  if (!img) throw new Error('导出失败: 无法生成底图')
  return canvasToBlob(img.canvas, opts.type ?? 'png', opts.quality ?? 0.92)
}
async function downloadImage(opts: ImageExportOptions = {}): Promise<void> {
  const blob = await exportImage(opts)
  const ext = opts.type === 'jpeg' ? 'jpg' : opts.type === 'webp' ? 'webp' : 'png'
  downloadBlob(blob, opts.fileName ?? `${baseName()}.${ext}`)
}

/** 导出为 PDF Blob(每个目标表分页;需可选依赖 jspdf)。未显式指定的页面参数取自工作表 pageSetup。 */
async function exportPdf(opts: PdfExportOptions = {}): Promise<Blob> {
  const targets = resolveTargets(opts.target)
  if (!targets.length) throw new Error('无可导出的工作表')
  const eff: PdfExportOptions = { ...pageSetupDefaults(targets[0]), ...opts }
  if (eff.vector) {
    const vs = (await Promise.all(targets.map((i) => buildVectorSheet(i, eff)))).filter(Boolean) as VectorSheet[]
    return exportToVectorPdf(vs, eff)
  }
  const images = (await Promise.all(targets.map((i) => buildSheetImage(i, eff, true)))).filter(Boolean) as ExportSheetImage[]
  return exportToPdf(images, eff)
}
async function downloadPdf(opts: PdfExportOptions = {}): Promise<void> {
  const blob = await exportPdf(opts)
  downloadBlob(blob, opts.fileName ?? `${baseName()}.pdf`)
}

/** 打开系统打印(可在对话框另存为 PDF)。页面参数同样默认取自 pageSetup。 */
async function print(opts: PrintOptions = {}): Promise<void> {
  const targets = resolveTargets(opts.target)
  if (!targets.length) return
  const eff: PrintOptions = { ...pageSetupDefaults(targets[0]), ...opts }
  const images = (await Promise.all(targets.map((i) => buildSheetImage(i, eff, true)))).filter(Boolean) as ExportSheetImage[]
  printSheets(images, { ...eff, title: eff.title ?? baseName() })
}

/** 工具栏触发 PDF: 捕获 jspdf 缺失等错误,给用户可读提示而非静默失败 */
async function onExportPdf() {
  try {
    await downloadPdf()
  } catch (e) {
    reportExportError(e)
  }
}
async function onExportPdfVector() {
  try {
    await downloadPdf({ vector: true })
  } catch (e) {
    reportExportError(e)
  }
}
function reportExportError(e: unknown) {
  const msg = (e as Error)?.message || String(e)
  emit('error', msg)
  if (typeof window !== 'undefined' && window.alert) window.alert(msg)
}

// ---- 查找 ----
const findOpen = ref(false)
const findQuery = ref('')
const findMatchCase = ref(false)
const findWholeCell = ref(false)
const findHits = shallowRef<{ row: number; col: number }[]>([])
const findIndex = ref(-1)

/** 重新计算命中并应用(query/选项变化时) */
function recomputeFind() {
  const r = renderer.value
  if (!r || !findQuery.value) {
    findHits.value = []
    findIndex.value = -1
    r?.setFind([], -1)
    doRender()
    return
  }
  const hits = r.searchCells(findQuery.value, { matchCase: findMatchCase.value, wholeCell: findWholeCell.value })
  findHits.value = hits
  findIndex.value = hits.length ? 0 : -1
  applyFind()
}

/** 把当前命中应用到渲染器 + 移动选区/滚动到视图 */
function applyFind() {
  const r = renderer.value
  if (!r) return
  r.setFind(findHits.value, findIndex.value)
  const hit = findHits.value[findIndex.value]
  if (hit) {
    selMode.value = 'range'
    selAnchor.value = { row: hit.row, col: hit.col }
    selActive.value = { row: hit.row, col: hit.col }
    scrollActiveIntoView()
  }
  doRender()
}

function findNext() {
  if (!findHits.value.length) return
  findIndex.value = (findIndex.value + 1) % findHits.value.length
  applyFind()
}
function findPrev() {
  if (!findHits.value.length) return
  findIndex.value = (findIndex.value - 1 + findHits.value.length) % findHits.value.length
  applyFind()
}
function openFind() {
  findOpen.value = true
}
function closeFind() {
  findOpen.value = false
  findQuery.value = ''
  findHits.value = []
  findIndex.value = -1
  renderer.value?.setFind([], -1)
  doRender()
  scrollerEl.value?.focus()
}
watch([findQuery, findMatchCase, findWholeCell], recomputeFind)

/** 根容器捕获 Ctrl/Cmd+F → 打开查找(替代浏览器原生查找) */
function onRootKeydown(e: KeyboardEvent) {
  if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) {
    e.preventDefault()
    openFind()
  }
}

// ---- 自动筛选 ----
// 列 → 允许值集合(缺省=该列未筛选);仅作用于当前表
const filterState = new Map<number, Set<string>>()
// 行 → 原始 hidden(首次筛选前快照,清除时恢复)
const filterOrigHidden = new Map<number, boolean>()
const filterPopup = ref<{ col: number; values: string[]; selected: string[]; x: number; y: number } | null>(null)

const BLANK = '(空白)'

/** 筛选数据区底行: 正常用 af.bottom;若 af 只含表头(bottom===top)则延伸到数据末行 */
function filterDataBottom(): number {
  const s = sheet.value!
  const af = s.autoFilterRange!
  return af.bottom > af.top ? af.bottom : s.dimension.rows - 1
}

/** 某列(自动筛选数据区)的去重值,数值/中文自然排序 */
function distinctColumnValues(col: number): string[] {
  const r = renderer.value
  const s = sheet.value
  if (!r || !s?.autoFilterRange) return []
  const af = s.autoFilterRange
  const set = new Set<string>()
  for (let row = af.top + 1; row <= filterDataBottom(); row++) set.add(r.cellText(row, col) || BLANK)
  return [...set].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
}

/** 重算筛选导致的隐藏行并应用到模型(行隐藏机制 → 几何归零) */
function applyFilters() {
  const r = renderer.value
  const s = sheet.value
  if (!r || !s?.autoFilterRange) return
  const af = s.autoFilterRange
  const bottom = filterDataBottom()
  if (!filterOrigHidden.size) {
    for (let row = af.top + 1; row <= bottom; row++) filterOrigHidden.set(row, s.rows.get(row)?.hidden ?? false)
  }
  for (let row = af.top + 1; row <= bottom; row++) {
    const orig = filterOrigHidden.get(row) ?? false
    let excluded = false
    for (const [col, allowed] of filterState) {
      if (!allowed.has(r.cellText(row, col) || BLANK)) {
        excluded = true
        break
      }
    }
    const hidden = orig || excluded
    const info = s.rows.get(row)
    if (info) info.hidden = hidden
    else if (hidden) s.rows.set(row, { height: s.defaultRowHeight, hidden: true })
  }
  r.setFilteredCols(new Set(filterState.keys()))
  r.rebuildMetrics()
  controller?.refreshContentSize()
  clearSelection()
  doRender()
}

function openFilterPopup(col: number) {
  const r = renderer.value
  const s = sheet.value
  if (!r || !s?.autoFilterRange) return
  const rect = r.cellScreenRect(view.value, s.autoFilterRange.top, col)
  let x = rect.x
  let y = rect.y + rect.h
  if (x + 228 > view.value.width) x = Math.max(0, view.value.width - 232)
  if (y + 320 > view.value.height) y = Math.max(0, rect.y - 320)
  filterPopup.value = {
    col,
    values: distinctColumnValues(col),
    selected: filterState.has(col) ? [...filterState.get(col)!] : [],
    x,
    y,
  }
}
function onFilterApply(checked: string[]) {
  const pop = filterPopup.value
  if (!pop) return
  const all = distinctColumnValues(pop.col)
  if (checked.length >= all.length) filterState.delete(pop.col) // 全选 = 取消该列筛选
  else filterState.set(pop.col, new Set(checked)) // 子集(含空集 = 全隐藏)
  filterPopup.value = null
  applyFilters()
}
function onFilterClear() {
  const pop = filterPopup.value
  if (!pop) return
  filterState.delete(pop.col)
  filterPopup.value = null
  applyFilters()
}
/** 离开某表时恢复其被筛选隐藏的行,清空筛选态 */
function resetFilterFor(idx: number | undefined) {
  if (idx == null) return
  const s = workbook.value?.sheets[idx]
  if (s && filterOrigHidden.size) {
    for (const [row, orig] of filterOrigHidden) {
      const info = s.rows.get(row)
      if (info) info.hidden = orig
    }
  }
  filterOrigHidden.clear()
  filterState.clear()
  filterPopup.value = null
}

/** 清除当前表全部筛选 */
function clearAllFilters() {
  if (!filterState.size) return
  filterState.clear()
  applyFilters()
}

/** 在活动单元格处冻结 / 取消冻结 */
function toggleFreeze() {
  const s = sheet.value
  const r = renderer.value
  if (!s || !r) return
  const fz = s.freeze
  if (fz.frozenRows || fz.frozenCols) {
    s.freeze = { frozenRows: 0, frozenCols: 0 }
  } else {
    const c = selActive.value
    s.freeze = { frozenRows: c ? c.row : 1, frozenCols: c ? c.col : 0 }
  }
  r.rebuildMetrics()
  controller?.refreshContentSize()
  doRender()
}

/** 工具栏「筛选」: 切换自动筛选。无则按选区(或整张已用区)新建,使下拉按钮出现。 */
function toggleAutoFilter() {
  const s = sheet.value
  const r = renderer.value
  if (!s || !r) return
  if (s.autoFilterRange) {
    resetFilterFor(activeSheet.value) // 恢复筛选隐藏的行 + 清状态
    s.autoFilterRange = undefined
  } else {
    const sel = selection.value
    const multi = sel && !(sel.top === sel.bottom && sel.left === sel.right)
    s.autoFilterRange = multi
      ? { ...sel! }
      : { top: 0, left: 0, bottom: Math.max(0, s.dimension.rows - 1), right: Math.max(0, s.dimension.cols - 1) }
  }
  r.setFilteredCols(new Set())
  r.rebuildMetrics()
  controller?.refreshContentSize()
  doRender()
}

// ---- 导出设置对话框 ----
const exportDialogOpen = ref(false)
/** 把对话框配置映射成各导出方法的入参并执行 */
async function onDialogExport(cfg: ExportConfig) {
  exportDialogOpen.value = false
  const target = cfg.scope === 'all' ? 'all' : 'active'
  const range = cfg.scope === 'selection' ? selection.value ?? undefined : undefined
  const common = {
    target: target as 'all' | 'active',
    range,
    scale: cfg.scale,
    includeHeaders: cfg.includeHeaders,
    gridlines: cfg.gridlines,
  }
  // 'auto' 表示沿用工作表 pageSetup,故不传(让 pageSetupDefaults 生效)
  const page = {
    ...(cfg.format !== 'auto' ? { format: cfg.format } : {}),
    ...(cfg.orientation !== 'auto' ? { orientation: cfg.orientation } : {}),
    fitToWidth: cfg.fitToWidth,
  }
  try {
    if (cfg.action === 'png') await downloadImage(common)
    else if (cfg.action === 'pdf') await downloadPdf({ ...common, ...page, vector: cfg.pdfVector })
    else await print({ ...common, ...page })
  } catch (e) {
    reportExportError(e)
  }
}

// ---------------- 命令式 API ----------------
function programmaticSetSelection(range: MergeRange) {
  selMode.value = 'range'
  selAnchor.value = { row: range.top, col: range.left }
  selActive.value = { row: range.bottom, col: range.right }
  doRender()
}

const viewerApi: ViewerApi = {
  load: (src: ExcelSource) => load(src, effectiveTransform),
  getWorkbook: () => workbook.value,
  getActiveSheet: () => activeSheet.value,
  setActiveSheet: (i: number) => {
    if (workbook.value?.sheets[i]) activeSheet.value = i
  },
  getSelection: () => selection.value,
  setSelection: programmaticSetSelection,
  rectOf,
  rectOfRange,
  redraw: () => doRender(),
  exportImage,
  downloadImage,
  exportPdf,
  downloadPdf,
  print,
  // ---- 数据读取(委托独立函数,自动绑 date1904 + 默认当前表) ----
  getCellValue: (row, col, si) => {
    const s = dataSheet(si)
    return s ? getCellValue(s, row, col) : null
  },
  getCellText: (row, col, si) => {
    const s = dataSheet(si)
    return s ? getCellText(s, row, col, workbook.value?.date1904 ?? false) : ''
  },
  getSheetData: (opts, si) => {
    const s = dataSheet(si)
    return s ? getSheetData(s, withDate1904(opts)) : []
  },
  getSheetJSON: (opts, si) => {
    const s = dataSheet(si)
    return s ? sheetToJSON(s, withDate1904(opts)) : []
  },
  getRangeData: (range, opts, si) => {
    const s = dataSheet(si)
    return s ? getRangeData(s, range, withDate1904(opts)) : []
  },
}
/** 取用于读数据的 sheet(缺省=当前活动表) */
function dataSheet(sheetIndex?: number): SheetModel | null {
  const wb = workbook.value
  if (!wb) return null
  return wb.sheets[sheetIndex ?? activeSheet.value] ?? null
}
function withDate1904<T extends ReadOptions>(opts?: T): T {
  return { ...(opts as T), date1904: workbook.value?.date1904 ?? false }
}
defineExpose(viewerApi)

// ---------------- 操作工具栏(可配置 + 可插件) ----------------
const I = (name: string) => TOOLBAR_ICONS[name]
function bi(o: Partial<ResolvedToolbarItem> & { id: string }): ResolvedToolbarItem {
  return { kind: 'builtin', ...o }
}
function builtinTool(id: string): ResolvedToolbarItem | null {
  switch (id) {
    case 'find':
      return bi({
        id,
        iconSvg: I('find'),
        label: '查找',
        title: '查找 (Ctrl+F)',
        active: findOpen.value,
        onClick: () => (findOpen.value ? closeFind() : openFind()),
      })
    case 'filter':
      return bi({
        id,
        iconSvg: I('filter'),
        label: '筛选',
        title: '切换自动筛选',
        active: !!sheet.value?.autoFilterRange,
        onClick: toggleAutoFilter,
      })
    case 'clear-filter':
      return bi({
        id,
        iconSvg: I('clear-filter'),
        label: '清除筛选',
        title: '清除当前表全部筛选',
        disabled: filterState.size === 0,
        onClick: clearAllFilters,
      })
    case 'copy':
      return bi({
        id,
        iconSvg: I('copy'),
        label: '复制',
        title: '复制选区 (Ctrl+C)',
        disabled: !selection.value,
        onClick: () => void copySelection(),
      })
    case 'freeze': {
      const fz = sheet.value?.freeze
      return bi({
        id,
        iconSvg: I('freeze'),
        label: '冻结',
        title: '冻结/取消冻结(在活动单元格)',
        active: !!(fz && (fz.frozenRows || fz.frozenCols)),
        onClick: toggleFreeze,
      })
    }
    case 'export':
      return bi({
        id,
        iconSvg: I('export'),
        label: '导出',
        title: '导出 / 打印',
        items: [
          bi({ id: 'export-png', label: '导出为图片 (PNG)', onClick: () => void downloadImage() }),
          bi({ id: 'export-pdf', label: '导出为 PDF (位图)', onClick: onExportPdf }),
          bi({ id: 'export-pdf-vector', label: '导出为 PDF (矢量·文字可选)', onClick: onExportPdfVector }),
          bi({ id: 'export-print', label: '打印…', onClick: () => void print() }),
          bi({ id: 'export-sep', type: 'separator' }),
          bi({ id: 'export-settings', label: '导出设置…', onClick: () => (exportDialogOpen.value = true) }),
        ],
      })
    case 'zoom':
      return bi({
        id,
        iconSvg: I('zoom'),
        label: Math.round(zoom.value * 100) + '%',
        title: '缩放',
        items: [50, 75, 100, 125, 150, 200].map((p) =>
          bi({ id: 'zoom-' + p, label: p + '%', active: Math.round(zoom.value * 100) === p, onClick: () => (zoom.value = p / 100) }),
        ),
      })
    default:
      return null // 'sort' 等待实现
  }
}

/** 把外来 ToolbarItem(自定义/插件)解析成 ResolvedToolbarItem(递归子菜单) */
function resolveItem(it: ToolbarItem, kind: 'custom' | 'plugin'): ResolvedToolbarItem {
  return {
    id: it.id,
    type: it.type,
    iconSvg: it.iconSvg,
    icon: it.icon,
    label: it.label,
    title: it.title,
    active: !!it.active?.(viewerApi),
    disabled: !!it.disabled?.(viewerApi),
    onClick: it.onClick ? () => it.onClick!(viewerApi) : undefined,
    items: it.items?.map((sub) => resolveItem(sub, kind)),
    kind,
  }
}

const resolvedToolbar = computed<ResolvedToolbarItem[]>(() => {
  void renderTick.value // 选区/状态变化时重算 active/disabled/label
  if (props.toolbar === false) return []
  const entries: Array<string | ToolbarItem> = Array.isArray(props.toolbar) ? props.toolbar : ['find', 'filter']
  const out: ResolvedToolbarItem[] = []
  for (const e of entries) {
    if (typeof e === 'string') {
      if (e === 'separator' || e === '|') out.push({ id: 'sep-' + out.length, type: 'separator', kind: 'builtin' })
      else {
        const b = builtinTool(e)
        if (b) out.push(b)
      }
    } else {
      out.push(resolveItem(e, 'custom'))
    }
  }
  for (const p of normalizedPlugins.value) {
    for (const it of p.toolbar ?? []) out.push(resolveItem(it, 'plugin'))
  }
  return out
})
const showActionBar = computed(() => props.toolbar !== false && resolvedToolbar.value.length > 0)

// ---------------- 插件运行时 ----------------
const pluginHandlers = new Map<PluginEvent, Set<(p: any) => void>>()
let pluginCleanups: Array<() => void> = []

/** 派发交互事件: 既 emit 给模板,也通知插件 */
function fire(event: PluginEvent, payload: any) {
  emit(event as any, payload)
  pluginHandlers.get(event)?.forEach((h) => h(payload))
}

/** (重新)初始化插件: 清理旧的 → 注册 events → 跑 setup 收集清理 */
function initPlugins() {
  pluginCleanups.forEach((fn) => fn())
  pluginCleanups = []
  pluginHandlers.clear()
  const register = (event: PluginEvent, fn: (p: any) => void) => {
    let set = pluginHandlers.get(event)
    if (!set) pluginHandlers.set(event, (set = new Set()))
    set.add(fn)
  }
  const ctx: ExcelPluginContext = { viewer: viewerApi, on: register, redraw: () => doRender() }
  for (const p of normalizedPlugins.value) {
    if (p.events) for (const [ev, fn] of Object.entries(p.events)) if (fn) register(ev as PluginEvent, fn)
    const cleanup = p.setup?.(ctx)
    if (typeof cleanup === 'function') pluginCleanups.push(cleanup)
  }
}

/** 渲染插件 overlay(读 renderTick 以随滚动/缩放重渲) */
const PluginOverlays = defineComponent({
  name: 'PluginOverlays',
  setup() {
    return () => {
      const ctx: OverlayContext = {
        rectOf,
        rectOfRange,
        tick: renderTick.value,
        workbook: workbook.value,
      }
      return normalizedPlugins.value.filter((p) => p.overlay).map((p) => p.overlay!(ctx))
    }
  },
})
</script>

<template>
  <div class="excel-viewer" ref="rootEl" @keydown="onRootKeydown">
    <slot
      v-if="workbook"
      name="header"
      :workbook="workbook"
      :zoom="zoom"
      :set-zoom="(z: number) => (zoom = z)"
      :download-image="downloadImage"
      :download-pdf="downloadPdf"
      :print="print"
    >
      <ViewerToolbar
        :file-name="fileName"
        :sheet-count="workbook.sheets.filter((s) => s.state === 'visible').length"
        :zoom="zoom"
        @update:zoom="zoom = $event"
        @export-image="downloadImage()"
        @export-pdf="onExportPdf"
        @export-pdf-vector="onExportPdfVector"
        @print="print()"
        @open-settings="exportDialogOpen = true"
      />
    </slot>

    <!-- 可配置/可插件 操作工具栏(查找/筛选/排序 + 插件项) -->
    <slot v-if="workbook && showActionBar" name="toolbar" :items="resolvedToolbar">
      <ActionToolbar :items="resolvedToolbar" />
    </slot>

    <div v-if="workbook" class="formula-bar">
      <span class="addr">{{ activeCellAddr || '—' }}</span>
      <span class="fx">fx</span>
      <span class="content" :title="formulaBarText">{{ formulaBarText }}</span>
    </div>

    <div class="render-area" ref="renderAreaEl">
      <canvas ref="canvasEl" class="grid-canvas" />
      <!-- 叠加层四象限(DOM 顺序=层级: 主区在下、冻结角在上) -->
      <div class="ov" ref="ovMain" />
      <div class="ov" ref="ovFCol" />
      <div class="ov" ref="ovFRow" />
      <div class="ov" ref="ovCorner" />
      <div
        class="scroller"
        ref="scrollerEl"
        tabindex="0"
        @scroll="onScroll"
        @mousedown="onMouseDown"
        @mousemove="onMouseMove"
        @mouseup="onMouseUp"
        @mouseleave="onMouseLeave"
        @dblclick="onDblClick"
        @keydown="onKeyDown"
      >
        <div class="spacer" ref="spacerEl" />
      </div>

      <FindBar
        v-if="findOpen && workbook"
        :query="findQuery"
        :match-count="findHits.length"
        :current="findIndex"
        :match-case="findMatchCase"
        :whole-cell="findWholeCell"
        @update:query="findQuery = $event"
        @update:match-case="findMatchCase = $event"
        @update:whole-cell="findWholeCell = $event"
        @next="findNext"
        @prev="findPrev"
        @close="closeFind"
      />

      <FilterPopup
        v-if="filterPopup"
        :values="filterPopup.values"
        :selected="filterPopup.selected"
        :x="filterPopup.x"
        :y="filterPopup.y"
        @apply="onFilterApply"
        @clear="onFilterClear"
        @close="filterPopup = null"
      />

      <!-- 分层 UI: 消费方在格子上叠自己的组件,用 rectOf 定位、tick 触发跟随 -->
      <div class="ov-slot">
        <slot name="overlay" :rect-of="rectOf" :rect-of-range="rectOfRange" :tick="renderTick" />
        <PluginOverlays />
      </div>

      <div
        v-if="tooltip"
        class="cell-tooltip"
        :class="tooltip.kind"
        :style="{ left: tooltip.x + 'px', top: tooltip.y + 'px' }"
      >
        {{ tooltip.text }}
      </div>

      <div v-if="loading" class="state">
        <slot name="loading" :progress="progress" :label="progressLabel" :pct="progressPct">
          <div class="loader">
            <div class="loader-label">
              {{ progressLabel }}<span v-if="progressPct != null"> {{ progressPct }}%</span>
            </div>
            <div class="loader-track">
              <div
                v-if="progressPct != null"
                class="loader-fill"
                :style="{ width: progressPct + '%' }"
              />
              <div v-else class="loader-fill indeterminate" />
            </div>
          </div>
        </slot>
      </div>
      <div v-else-if="error" class="state error">
        <slot name="error" :error="error">解析失败：{{ error }}</slot>
      </div>
      <div v-else-if="!workbook" class="state hint">
        <slot name="empty">拖入或选择一个 .xlsx 文件</slot>
      </div>
    </div>

    <div v-if="workbook" class="status-bar">
      <slot name="statusbar" :stats="stats" :range="selRangeLabel || activeCellAddr">
        <span class="sel">{{ selRangeLabel || activeCellAddr }}</span>
        <div class="grow" />
        <template v-if="stats && stats.numCount > 0">
          <span>计数 {{ stats.count }}</span>
          <span>求和 {{ fmtNum(stats.sum) }}</span>
          <span>平均 {{ fmtNum(stats.avg) }}</span>
          <span>最大 {{ fmtNum(stats.max) }}</span>
          <span>最小 {{ fmtNum(stats.min) }}</span>
        </template>
        <span v-else-if="stats && stats.count > 0">计数 {{ stats.count }}</span>
      </slot>
    </div>

    <SheetTabs
      v-if="workbook"
      :workbook="workbook as WorkbookModel"
      :active="activeSheet"
      @select="activeSheet = $event"
    />

    <ExportDialog
      v-if="exportDialogOpen && workbook"
      :selection="selection"
      :sheet-count="workbook.sheets.filter((s) => s.state === 'visible').length"
      @close="exportDialogOpen = false"
      @export="onDialogExport"
    />
  </div>
</template>

<style scoped>
.excel-viewer {
  display: flex;
  flex-direction: column;
  height: 100%;
  width: 100%;
  background: #fff;
  overflow: hidden;
}
.render-area {
  position: relative;
  flex: 1 1 auto;
  overflow: hidden;
  min-height: 0;
}
.grid-canvas {
  position: absolute;
  inset: 0;
  z-index: 1;
}
/* 叠加层四象限: 各自裁剪,层级靠 DOM 顺序(主区先=底,冻结角后=顶) */
.ov {
  position: absolute;
  z-index: 2;
  overflow: hidden;
  pointer-events: none;
}
/* 分层 UI slot: 叠在最上(滚动层之上),容器不吃事件,子元素可 pointer-events:auto 接收交互 */
.ov-slot {
  position: absolute;
  inset: 0;
  z-index: 4;
  overflow: hidden;
  pointer-events: none;
}
.ov-slot :deep(*) {
  pointer-events: auto;
}
/* 滚动条层: 透明、置顶、提供原生滚动条 + 接收鼠标交互 */
.scroller {
  position: absolute;
  inset: 0;
  z-index: 3;
  overflow: auto;
  cursor: cell;
  outline: none;
}
.spacer {
  pointer-events: none;
}
/* 公式栏 */
.formula-bar {
  display: flex;
  align-items: center;
  height: 28px;
  flex: 0 0 auto;
  border-bottom: 1px solid #e2e4e7;
  background: #fff;
  font-size: 13px;
}
.formula-bar .addr {
  width: 72px;
  text-align: center;
  border-right: 1px solid #e2e4e7;
  color: #444;
  font-weight: 600;
  flex: 0 0 auto;
}
.formula-bar .fx {
  width: 34px;
  text-align: center;
  color: #999;
  font-style: italic;
  flex: 0 0 auto;
}
.formula-bar .content {
  flex: 1;
  padding: 0 8px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  color: #222;
  font-family: Consolas, 'Courier New', monospace;
}
/* 状态栏 */
.status-bar {
  display: flex;
  align-items: center;
  gap: 16px;
  height: 24px;
  flex: 0 0 auto;
  padding: 0 12px;
  border-top: 1px solid #e2e4e7;
  background: #fbfbfb;
  font-size: 12px;
  color: #555;
}
.status-bar .grow {
  flex: 1;
}
.status-bar .sel {
  color: #888;
}
/* 裁切文本悬停提示 */
.cell-tooltip {
  position: absolute;
  z-index: 6;
  max-width: 380px;
  padding: 5px 9px;
  background: #2b2f33;
  color: #fff;
  font-size: 12px;
  line-height: 1.4;
  border-radius: 4px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.25);
  pointer-events: none;
  white-space: pre-wrap;
  word-break: break-all;
}
/* 批注样式: 仿 Excel 便签(浅黄底、深色字) */
.cell-tooltip.comment {
  background: #fdfcdc;
  color: #333;
  border: 1px solid #d9d27e;
  box-shadow: 0 2px 10px rgba(0, 0, 0, 0.18);
  max-width: 300px;
}
.state {
  position: absolute;
  inset: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 5;
  background: #fff;
  color: #888;
  font-size: 14px;
}
.state.error { color: #c0392b; }
.state.hint { color: #aaa; }
/* 分阶段加载进度 */
.loader {
  width: 260px;
  max-width: 60%;
}
.loader-label {
  font-size: 13px;
  color: #555;
  margin-bottom: 8px;
  text-align: center;
}
.loader-track {
  height: 6px;
  border-radius: 3px;
  background: #eceef0;
  overflow: hidden;
}
.loader-fill {
  height: 100%;
  background: #21a366;
  border-radius: 3px;
  transition: width 0.15s ease;
}
/* exceljs 黑盒阶段: 不确定态脉冲(诚实地表示"在动但拿不到 %") */
.loader-fill.indeterminate {
  width: 40%;
  transition: none;
  animation: loader-pulse 1.1s ease-in-out infinite;
}
@keyframes loader-pulse {
  0% { margin-left: -40%; }
  100% { margin-left: 100%; }
}
</style>
