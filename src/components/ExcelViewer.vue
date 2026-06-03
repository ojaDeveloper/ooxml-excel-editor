<script setup lang="ts">
import { computed, defineComponent, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, watch } from 'vue'
import type * as EChartsNS from 'echarts'
import type { ExcelPlugin, ExcelPluginContext, OverlayContext, PluginEvent, ViewerApi } from '@/core/plugin'
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
import type { ViewerTheme } from '@/core/render/theme'
import { useExcelDocument } from '@/composables/useExcelDocument'
import { CanvasRenderer, type ViewState } from '@/core/render/canvas-renderer'
import { colIndexToLetters } from '@/core/layout/grid-metrics'
import { anchorRect } from '@/core/overlay/anchor'
import { chartToOption } from '@/core/overlay/chart-mapper'
import { revokeImages } from '@/core/finalize'
import ViewerToolbar from './ViewerToolbar.vue'
import SheetTabs from './SheetTabs.vue'

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
  }>(),
  { openLinks: true },
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

const contentSize = ref({ w: 0, h: 0 })

// ---------------- 图表实例管理 ----------------
// echarts 按需加载: 只有 sheet 含图表时才动态 import，省掉无图表文件的 ~1MB 首包
let echartsMod: typeof EChartsNS | null = null
async function loadECharts(): Promise<typeof EChartsNS> {
  if (!echartsMod) echartsMod = await import('echarts')
  return echartsMod
}

type Quad = 'main' | 'frow' | 'fcol' | 'corner'
interface ChartInstance {
  el: HTMLDivElement
  inst: EChartsNS.ECharts
  spec: ChartSpec
  quad: Quad
  lastW: number // 上次尺寸,只有变化才 resize(避免滚动时每帧 resize)
  lastH: number
}
let chartInstances: ChartInstance[] = []
interface ImageEl {
  el: HTMLImageElement
  anchorIdx: number
  quad: Quad
}
let imageEls: ImageEl[] = []
// echarts 缺失时的占位框(仍按图表锚点定位，给出友好提示)
interface ChartPlaceholder {
  el: HTMLDivElement
  spec: ChartSpec
  quad: Quad
}
let chartPlaceholders: ChartPlaceholder[] = []
interface ShapeEl {
  el: HTMLDivElement
  shapeIdx: number
  quad: Quad
}
let shapeEls: ShapeEl[] = []

/** 锚点落在哪个象限(冻结行/列内的钉住,其余随滚动) */
function quadrantOf(anchor: ImageAnchor): Quad {
  const fz = renderer.value!.freezeGeometry
  const top = anchor.from.row < fz.frozenRows
  const left = anchor.from.col < fz.frozenCols
  if (top && left) return 'corner'
  if (top) return 'frow'
  if (left) return 'fcol'
  return 'main'
}
function ovContainer(q: Quad): HTMLElement | null {
  return q === 'main' ? ovMain.value : q === 'frow' ? ovFRow.value : q === 'fcol' ? ovFCol.value : ovCorner.value
}

function disposeOverlays() {
  for (const c of chartInstances) {
    c.inst.dispose()
    c.el.remove()
  }
  chartInstances = []
  for (const p of chartPlaceholders) p.el.remove()
  chartPlaceholders = []
  for (const sh of shapeEls) sh.el.remove()
  shapeEls = []
  for (const im of imageEls) im.el.remove()
  imageEls = []
}

async function buildOverlays() {
  disposeOverlays()
  const s = sheet.value
  if (!s || !ovMain.value) return

  for (let i = 0; i < s.images.length; i++) {
    const img = s.images[i]
    const quad = quadrantOf(img)
    const el = document.createElement('img')
    el.src = img.src
    el.draggable = false
    el.style.position = 'absolute'
    el.style.objectFit = 'fill'
    el.style.pointerEvents = 'none'
    ovContainer(quad)?.appendChild(el)
    imageEls.push({ el, anchorIdx: i, quad })
  }

  // 形状 / 文本框
  for (let i = 0; i < s.shapes.length; i++) {
    const shape = s.shapes[i]
    const quad = quadrantOf(shape.anchor)
    const el = document.createElement('div')
    el.style.position = 'absolute'
    el.style.boxSizing = 'border-box'
    el.style.pointerEvents = 'none'
    el.style.overflow = 'hidden'
    el.style.display = 'flex'
    el.style.padding = '3px 5px'
    el.style.alignItems = 'center'
    el.style.justifyContent = shape.align === 'center' ? 'center' : shape.align === 'right' ? 'flex-end' : 'flex-start'
    el.style.whiteSpace = 'pre-wrap'
    el.style.wordBreak = 'break-word'
    el.style.lineHeight = '1.2'
    if (shape.fillColor) el.style.background = shape.fillColor
    if (shape.lineColor) el.style.border = `1px solid ${shape.lineColor}`
    if (shape.geom === 'roundRect') el.style.borderRadius = '8px'
    else if (shape.geom === 'ellipse') el.style.borderRadius = '50%'
    if (shape.textColor) el.style.color = shape.textColor
    if (shape.bold) el.style.fontWeight = 'bold'
    el.style.textAlign = shape.align ?? 'left'
    el.textContent = shape.text ?? ''
    ovContainer(quad)?.appendChild(el)
    shapeEls.push({ el, shapeIdx: i, quad })
  }

  if (s.charts.length) {
    let echarts: typeof EChartsNS | null = null
    try {
      echarts = await loadECharts()
    } catch (e) {
      // echarts 是可选 peer 依赖，宿主未安装时会走到这里
      console.warn('[ooxml-preview] 未能加载 echarts，图表降级为占位提示。请安装依赖: npm i echarts', e)
    }
    // 异步加载期间可能已切表，确认当前仍是该 sheet 才挂
    if (sheet.value !== s) return

    for (const chart of s.charts) {
      const quad = quadrantOf(chart.anchor)
      if (echarts) {
        const el = document.createElement('div')
        el.style.cssText = chartBoxCss
        ovContainer(quad)?.appendChild(el)
        const inst = echarts.init(el)
        try {
          inst.setOption(chartToOption(chart))
        } catch (e) {
          console.warn('[ooxml-preview] 图表渲染失败:', e)
        }
        chartInstances.push({ el, inst, spec: chart, quad, lastW: 0, lastH: 0 })
      } else {
        // 降级: 在图表位置画一个友好占位框
        const el = document.createElement('div')
        el.style.cssText = chartBoxCss + placeholderCss
        el.innerHTML =
          '<div style="font-size:22px;line-height:1">📊</div>' +
          '<div style="margin-top:6px;font-weight:600">图表</div>' +
          '<div style="margin-top:2px;font-size:11px;color:#9aa4ae">渲染需安装 echarts 依赖</div>'
        ovContainer(quad)?.appendChild(el)
        chartPlaceholders.push({ el, spec: chart, quad })
      }
    }
  }
  positionOverlays()
}

const chartBoxCss =
  'position:absolute;background:rgba(255,255,255,0.96);border:1px solid #e2e4e7;box-shadow:0 1px 4px rgba(0,0,0,0.08);'
const placeholderCss =
  'display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;color:#6b7785;border-style:dashed;box-sizing:border-box;padding:8px;overflow:hidden;'

function positionOverlays() {
  const r = renderer.value
  const s = sheet.value
  if (!r || !s) return
  const hw = r.metrics.rowHeaderWidth
  const hh = r.metrics.colHeaderHeight
  const fz = r.freezeGeometry
  const fw = fz.frozenWidth
  const fh = fz.frozenHeight
  const bodyW = Math.max(0, view.value.width - hw - fw)
  const bodyH = Math.max(0, view.value.height - hh - fh)

  // 四象限容器各自裁到对应区域(冻结角/行/列/主区)
  setQuad(ovCorner.value, hw, hh, fw, fh)
  setQuad(ovFRow.value, hw + fw, hh, bodyW, fh)
  setQuad(ovFCol.value, hw, hh + fh, fw, bodyH)
  setQuad(ovMain.value, hw + fw, hh + fh, bodyW, bodyH)

  for (const im of imageEls) {
    placeInQuad(im.el, anchorRect(r.metrics, s.images[im.anchorIdx]), im.quad, fw, fh)
  }
  const shapeFont = Math.round(11 * (96 / 72) * r.metrics.zoom)
  for (const sh of shapeEls) {
    sh.el.style.fontSize = shapeFont + 'px'
    placeInQuad(sh.el, anchorRect(r.metrics, s.shapes[sh.shapeIdx].anchor), sh.quad, fw, fh)
  }
  for (const c of chartInstances) {
    const rect = anchorRect(r.metrics, c.spec.anchor)
    placeInQuad(c.el, rect, c.quad, fw, fh)
    // 只有尺寸真变(缩放)才 resize;滚动只改位置,不必 resize(很费)
    if (c.lastW !== rect.width || c.lastH !== rect.height) {
      c.inst.resize()
      c.lastW = rect.width
      c.lastH = rect.height
    }
  }
  for (const p of chartPlaceholders) {
    placeInQuad(p.el, anchorRect(r.metrics, p.spec.anchor), p.quad, fw, fh)
  }
}

function setQuad(el: HTMLElement | null, left: number, top: number, w: number, h: number) {
  if (!el) return
  el.style.left = left + 'px'
  el.style.top = top + 'px'
  el.style.width = Math.max(0, w) + 'px'
  el.style.height = Math.max(0, h) + 'px'
}

/** 把元素定位到所在象限容器内的相对坐标(冻结方向不减滚动量) */
function placeInQuad(
  el: HTMLElement,
  rect: { left: number; top: number; width: number; height: number },
  quad: Quad,
  fw: number,
  fh: number,
) {
  const sx = view.value.scrollX
  const sy = view.value.scrollY
  // 主区/冻结行的横向随滚动(容器从 fw 处起,故减 fw);冻结列/角横向固定
  const x = quad === 'main' || quad === 'frow' ? rect.left - fw - sx : rect.left
  // 主区/冻结列的纵向随滚动(容器从 fh 处起,故减 fh);冻结行/角纵向固定
  const y = quad === 'main' || quad === 'fcol' ? rect.top - fh - sy : rect.top
  el.style.left = x + 'px'
  el.style.top = y + 'px'
  el.style.width = rect.width + 'px'
  el.style.height = rect.height + 'px'
}

// ---------------- 渲染 ----------------
// overlay slot 用: 每次重绘 +1 → 作用域插槽重算 rectOf 位置(随滚动/缩放/切表跟随)
const renderTick = ref(0)
let rafId = 0
/** 把重绘合并到下一帧(滚动/拖选高频触发时,每帧最多画一次) */
function scheduleRender() {
  if (rafId) return
  rafId = requestAnimationFrame(() => {
    rafId = 0
    doRender()
  })
}
function doRender() {
  if (rafId) {
    cancelAnimationFrame(rafId)
    rafId = 0
  }
  const r = renderer.value
  if (!r) return
  view.value.zoom = zoom.value
  r.setSelection(selection.value)
  r.render(view.value)
  positionOverlays()
  renderTick.value++ // 通知 overlay slot 重算位置
}

function rebuildRenderer() {
  const s = sheet.value
  const wb = workbook.value
  const canvas = canvasEl.value
  if (!s || !wb || !canvas) return
  renderer.value = new CanvasRenderer(canvas, s, wb, zoom.value, {
    theme: effectiveTheme.value,
    cellStyle: hasCellStyleHook.value ? combinedCellStyle : undefined,
  })
  contentSize.value = { w: renderer.value.contentWidth, h: renderer.value.contentHeight }
  clearSelection() // 切表清空选区
  tooltip.value = null
  // 重置滚动
  if (scrollerEl.value) {
    scrollerEl.value.scrollLeft = 0
    scrollerEl.value.scrollTop = 0
  }
  view.value.scrollX = 0
  view.value.scrollY = 0
  measure()
  buildOverlays()
  doRender()
}

function measure() {
  const area = renderAreaEl.value
  if (!area) return
  view.value.width = area.clientWidth
  view.value.height = area.clientHeight
}

function onScroll() {
  const sc = scrollerEl.value
  if (!sc) return
  view.value.scrollX = sc.scrollLeft
  view.value.scrollY = sc.scrollTop
  tooltip.value = null
  scheduleRender()
}

let resizeObserver: ResizeObserver | null = null

onMounted(() => {
  initPlugins()
  if (props.src) load(props.src, effectiveTransform)
  resizeObserver = new ResizeObserver(() => {
    measure()
    doRender()
  })
  if (renderAreaEl.value) resizeObserver.observe(renderAreaEl.value)
})

onBeforeUnmount(() => {
  resizeObserver?.disconnect()
  if (rafId) cancelAnimationFrame(rafId)
  pluginCleanups.forEach((fn) => fn())
  disposeOverlays()
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

watch(activeSheet, async () => {
  await nextTick()
  rebuildRenderer()
})

watch(zoom, async (z) => {
  const r = renderer.value
  const sc = scrollerEl.value
  if (!r) return
  // 缩放前记录视口中心相对内容的比例，缩放后还原中心，避免"跳到左上角"
  const ratioX = sc && contentSize.value.w ? (sc.scrollLeft + sc.clientWidth / 2) / contentSize.value.w : 0
  const ratioY = sc && contentSize.value.h ? (sc.scrollTop + sc.clientHeight / 2) / contentSize.value.h : 0
  r.setZoom(z)
  contentSize.value = { w: r.contentWidth, h: r.contentHeight }
  await nextTick()
  if (sc) {
    sc.scrollLeft = Math.max(0, ratioX * contentSize.value.w - sc.clientWidth / 2)
    sc.scrollTop = Math.max(0, ratioY * contentSize.value.h - sc.clientHeight / 2)
    view.value.scrollX = sc.scrollLeft
    view.value.scrollY = sc.scrollTop
  }
  doRender()
})

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
      r.setColWidthPx(resizeTarget, resizeStartSize + (p.x - resizeStartPos))
      contentSize.value = { w: r.contentWidth, h: r.contentHeight }
      scheduleRender()
      return
    }
    if (dragMode === 'resize-row') {
      r.setRowHeightPx(resizeTarget, resizeStartSize + (p.y - resizeStartPos))
      contentSize.value = { w: r.contentWidth, h: r.contentHeight }
      scheduleRender()
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
    r.autoFitColumn(colHit.col)
  } else if (rowHit) {
    r.autoFitRow(rowHit.row)
  } else {
    // 非边界 → 双击单元格事件
    const cell = r.cellAtScreen(view.value, p.x, p.y)
    if (cell) fire('cell-dblclick', { row: cell.row, col: cell.col, text: r.cellText(cell.row, cell.col) })
    return
  }
  contentSize.value = { w: r.contentWidth, h: r.contentHeight }
  doRender()
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
  const r = renderer.value
  if (!r) return null
  return r.screenRectOfCell(view.value, row, col)
}
/** 区域当前屏幕矩形(左上到右下的并集) */
function rectOfRange(range: MergeRange): { x: number; y: number; w: number; h: number } | null {
  const r = renderer.value
  if (!r) return null
  const tl = r.screenRectOfCell(view.value, range.top, range.left)
  const br = r.screenRectOfCell(view.value, range.bottom, range.right)
  return { x: tl.x, y: tl.y, w: br.x + br.w - tl.x, h: br.y + br.h - tl.y }
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
}
defineExpose(viewerApi)

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
  <div class="excel-viewer" ref="rootEl">
    <slot v-if="workbook" name="toolbar" :workbook="workbook" :zoom="zoom" :set-zoom="(z: number) => (zoom = z)">
      <ViewerToolbar
        :file-name="fileName"
        :sheet-count="workbook.sheets.filter((s) => s.state === 'visible').length"
        :zoom="zoom"
        @update:zoom="zoom = $event"
      />
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
        <div class="spacer" :style="{ width: contentSize.w + 'px', height: contentSize.h + 'px' }" />
      </div>

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
