<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, watch } from 'vue'
import type { ExcelPlugin, ExcelPluginContext, OverlayContext, PermissionDeniedPayload, PluginEvent, ToolbarItem, ViewerApi } from '@/core/plugin'
import { PluginOverlayHost } from '@/core/viewer/plugin-overlay'
import { loadArrayBuffer, type ExcelSource } from '@/core/loader'
import { jsonToWorkbook, isWorkbookModel, type JsonInput, type JsonLoadOptions } from '@/core/loader-json'
import { applyStyleTemplate } from '@/core/template/style-overlay'
import { detectFormat, finalizeImages } from '@/core/finalize'
import { parseInWorker } from '@/composables/worker-client'
import type {
  CellModel,
  CellStyleFn,
  CellStyleOverride,
  MergeRange,
  SheetModel,
  TransformModelFn,
  WorkbookModel,
} from '@/core/model/types'
import type { ExportProgress, ParseProgress } from '@/core/progress'
import type { ReadOptions } from '@/core/model/data-access'
import { getCellValue, getCellText, getSheetData, getRangeData, sheetToJSON } from '@/core/model/data-access'
import type { ViewerTheme } from '@/core/render/theme'
import { useExcelDocument } from '@/composables/useExcelDocument'
import { CanvasRenderer, type ViewState } from '@/core/render/canvas-renderer'
import { colIndexToLetters } from '@/core/layout/grid-metrics'
import { ViewerController, type ContextMenuBeforePayload, type ContextMenuShowPayload, type ContextMenuTransform, type TooltipState, type FindState } from '@/core/viewer/controller'
import type { EditableTarget, EditConfig } from '@/core/edit/types'
import type { PasteBehavior } from '@/core/edit/paste-behavior'
import { DEFAULT_PASTE_BEHAVIOR } from '@/core/edit/paste-behavior'
import type { FormulaEngineFactory } from '@/core/formula/engine'
import type { CellChangePayload, DimChangePayload, DirtyChangePayload, ImageChangePayload, StructChangePayload } from '@/core/edit/edit-controller'
import type { EditorResolver, CellEditorFactory } from '@/core/edit/editor-context'
import { revokeImages } from '@/core/finalize'
import type { ImageExportOptions, PdfExportOptions, PrintOptions } from '@/core/export'
import ViewerToolbar from './ViewerToolbar.vue'
import SheetTabs from './SheetTabs.vue'
import ExportDialog from './ExportDialog.vue'
import FindBar from './FindBar.vue'
import FilterPopup from './FilterPopup.vue'
import ActionToolbar from './ActionToolbar.vue'
import ExportProgressOverlay from './ExportProgressOverlay.vue'
import type { ExportConfig } from './export-types'
import type { ResolvedToolbarItem } from './toolbar-types'
import { TOOLBAR_ICONS } from './toolbar-icons'

const props = withDefaults(
  defineProps<{
    src?: ExcelSource
    /**
     * 直接喂 WorkbookModel 或 JsonInput(P3):绕过 parser,常用于"从后端拿 JSON 直接渲染"或
     * "前端构造好模型再渲染"。WorkbookModel-shape 直用,JsonInput(二维数组 / 对象数组 / `{sheets:[...]}`)
     * 走 `jsonToWorkbook` 自动构造。优先级:`workbook` > `src`(两者都给时取 `workbook`)。
     */
    workbook?: WorkbookModel | JsonInput
    /** JSON 直渲选项(`workbook` = JsonInput 时生效) */
    jsonOptions?: JsonLoadOptions
    /**
     * 渲染模板(P3 重设计 2026-06-08):一份 .xlsx 当**样式捐赠者** —— 模板的 styling
     * (styles / merges / 列宽 / 行高 / freeze / theme)套到无格式数据源上,模板的 raw 文字 / 占位符
     * / 图 / 图表 / 条件格式 全部丢弃。
     *
     * ⚠️ **只在数据源是 :workbook(JSON / 模型)时生效**;`:src`(.xlsx)数据源自带格式,
     * 给 `:templateFile` 会被忽略并 console.warn。
     *
     * 工具栏内置 `template` 项可在运行时切换/导入/清除,无需重新挂载。
     */
    templateFile?: ExcelSource
    /** 模板显示名(标题栏 `· 模板: xxx` 后缀);不给则取运行时 File.name */
    templateName?: string
    fileName?: string
    /** 外观主题(覆盖默认配色) */
    theme?: Partial<ViewerTheme>
    /** 数据钩子: 解析后改模型再渲染 */
    transformModel?: TransformModelFn
    /** 渲染钩子: 按单元格覆盖样式 */
    cellStyle?: CellStyleFn
    /** WPS 单元格内嵌图(DISPIMG)贴合方式:contain 等比(默认,与 WPS 渲染一致)/ fill 拉伸铺满 / cover 等比裁剪 */
    cellImageFit?: 'fill' | 'contain' | 'cover'
    /** 图片点击放大灯箱(默认 true):只读模式单击图放大、编辑模式右键「查看大图」;false 关闭 */
    imageLightbox?: boolean
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
    /** 编辑总开关:默认 false = 只读(行为不变)。开启后才能进入编辑(E0:闸门) */
    editable?: boolean
    /**
     * 透视表功能开关:默认 false = 关闭。开启后(还需 `editable`)工具栏 `pivot-table` 入口可见、
     * `createPivotTable`/`openPivotTableDialog` 等 API 生效、导出 .xlsx 回注真实 OOXML 透视表零件
     * (overlay 模式同时保留原文件透视表)。
     */
    pivotTable?: boolean
    /**
     * 条件格式编辑开关:默认 false = 关闭(只读渲染)。开启后(还需 `editable`)工具栏 `conditional-format`
     * 入口可见、`openConditionalFormatDialog`/`addConditionalRule` 等 API 生效、导出 .xlsx 回写条件格式
     * (overlay 模式保留原件未编辑规则原样,只增改用户改的)。
     */
    conditionalFormat?: boolean
    /** 按格只读判定:返回 true = 只读(cell 为空格时传 null) */
    cellReadOnly?: (cell: CellModel | null, pos: { row: number; col: number }) => boolean | void
    /** 只读区域(0-based 闭区间);命中即只读 */
    readOnlyRanges?: MergeRange[]
    /**
     * **可编辑白名单**(2026-06-08 新增) —— 设了就是白名单语义:默认只读,只有命中**任一**
     * target 的格才可编辑。4 种 target 形状:`{row,col}` 单格 / `{row}` 整行 / `{col}` 整列 /
     * `MergeRange` 矩形。可单值可数组,允许**不相邻**多个 target.
     *
     * `undefined`(不传)= 不启用白名单(老行为:默认全可编辑);`[]`(显式空数组)= 全只读.
     * 与 `readOnlyRanges` / `cellReadOnly` 叠加 — 白名单命中后仍可被它们二次"黑"掉.
     */
    editableTargets?: EditableTarget | EditableTarget[]
    /**
     * **严格尺寸闸门**(Phase B, 2026-06-08) —— 默认 `false`: setColumnWidth/setRowHeight/autoFit 仅受全局
     * `editable` 控制(老行为). 设 `true` + `editableTargets` 启用了 → 该列/行至少有 1 格在白名单内才能改尺寸.
     */
    strictDimensions?: boolean
    /**
     * **只读单元格视觉钩子** (Phase C, 2026-06-08):
     * - `false` (默认) = 无视觉差异 (老行为不变)
     * - `true` = 套内置默认 (浅灰底 `#f5f7fa`,跟工具栏一致)
     * - `CellStyleOverride` 对象 = 固定样式给所有只读格 (如黄底高亮)
     * - `CellStyleFn` 函数 = 按格自定义
     *
     * 跟 `editableTargets` 配合: 白名单未覆盖的格自动套此视觉, 用户一眼看出哪些可编辑.
     */
    readOnlyCellStyle?: boolean | CellStyleOverride | CellStyleFn
    /** 自定义单元格编辑器(按格返回工厂;覆盖插件 editor)。需 editable 开启 */
    editor?: EditorResolver
    /** 公式重算(E4):默认 false 沿用缓存值。开启后编辑公式/被引用格 → 依赖格自动重算。需 editable */
    recalc?: boolean
    /** 自定义/自研公式引擎工厂(可换引擎);不给则用默认 HyperFormula(需 npm i hyperformula) */
    formulaEngine?: FormulaEngineFactory
    /**
     * 粘贴行为(默认 = 覆盖式 1:1)。控制 Ctrl+V / 右键粘贴时源各方面如何落目标(覆盖/合并/仅值)。
     * 不传 = 默认;也可 `viewer.setPasteBehavior(cfg)` 运行时改、右键「选择性粘贴」逐次选预设。
     */
    pasteBehavior?: Partial<PasteBehavior>
    /** 粘贴撞只读格的内置提醒:'dialog'(默认,弹窗列出哪些格只读)/ 'toast'(气泡)/ 'none'(只发事件) */
    readOnlyPrompt?: 'dialog' | 'toast' | 'none'
    /**
     * 内置导出进度遮罩(P1.5):默认 `true` —— 调 `viewer.downloadPdf` / `downloadImage` / `downloadXlsx` /
     * `print` / 选区图片批量转换 时,壳自动建 `AbortController` + 接 `onProgress` →
     * 显示居中模态(stage 标签 + 进度条 + 取消按钮)。**关闭** `:export-progress="false"` 走纯回调
     * 路径(用户自己接 `opts.onProgress`/`opts.signal`)。**完全自渲染**用 `#export-progress` 插槽
     * (拿到 `{ state, busy, cancel }`)。
     */
    exportProgress?: boolean
    /**
     * 右键菜单(Plan C):
     * - `false` → 不弹内置菜单(`before-context-menu` / `context-menu` 事件仍触发,用户自渲染)
     * - 函数 `(ctx, items) => MenuItem[] | undefined` → 在内置 items 上加 / 减 / 重排;返 `undefined` 用原样
     * - 不传(默认)→ editable 时显示内置菜单,非 editable 走浏览器默认菜单
     * 与 `@before-context-menu` / `@context-menu` 事件叠加使用。
     */
    contextMenu?: boolean | ContextMenuTransform
  }>(),
  // toolbar/imageLightbox/exportProgress 默认 true;若不显式给默认,Vue 会把布尔型 prop 缺省判成 false
  // contextMenu 显式默认 undefined(否则 Vue 会把 `boolean | Function` prop 缺省判成 false,把"未传"误解读为"关掉内置")
  { openLinks: true, toolbar: true, imageLightbox: true, exportProgress: true, contextMenu: undefined as undefined },
)

const normalizedPlugins = computed<ExcelPlugin[]>(() => props.plugins ?? [])

// 合并各扩展点: 插件按数组顺序,组件 prop 最后覆盖
const effectiveTheme = computed(() =>
  Object.assign({}, ...normalizedPlugins.value.map((p) => p.theme || {}), props.theme || {}),
)
const hasCellStyleHook = computed(() => !!props.cellStyle || normalizedPlugins.value.some((p) => p.cellStyle))
function combinedCellStyle(cell: CellModel, pos: { row: number; col: number }, ctx?: import('@/core/model/types').CellStyleCtx) {
  let acc: Record<string, unknown> | undefined
  const apply = (fn?: CellStyleFn) => {
    const o = fn?.(cell, pos, ctx) // Phase C 2026-06-08: 透传 ctx.editable 给插件 / props cellStyle
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
// 编辑配置(E0:默认只读;开 editable + 可按格/区域只读)
const effectiveEditConfig = computed<EditConfig>(() => ({
  editable: props.editable,
  pivotTable: props.pivotTable,
  conditionalFormat: props.conditionalFormat,
  cellReadOnly: props.cellReadOnly,
  readOnlyRanges: props.readOnlyRanges,
  editableTargets: props.editableTargets,
  strictDimensions: props.strictDimensions,
  recalc: props.recalc,
  formulaEngine: props.formulaEngine,
  pasteBehavior: props.pasteBehavior,
  readOnlyPrompt: props.readOnlyPrompt,
}))
// 合并编辑器解析器(E2:组件 editor prop 优先,其次插件 editor 数组序首个非空)
function resolveEditor(cell: CellModel | null, pos: { row: number; col: number }): CellEditorFactory | void {
  const fromProp = props.editor?.(cell, pos)
  if (fromProp) return fromProp
  for (const p of normalizedPlugins.value) {
    const f = p.editor?.(cell, pos)
    if (f) return f
  }
}
const hasEditor = computed(() => !!props.editor || normalizedPlugins.value.some((p) => p.editor))

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
  /** 单元格变更(编辑/撤销/重做;含前后完整快照) */
  (e: 'cell-change', payload: CellChangePayload): void
  /** 进入编辑 */
  (e: 'edit-start', payload: unknown): void
  /** 提交编辑 */
  (e: 'edit-commit', payload: unknown): void
  /** 列宽/行高变更(拖拽/autofit/API/撤销重做;前后 px 尺寸) */
  (e: 'dim-change', payload: DimChangePayload): void
  /** 脏状态变更(有/无未保存修改) */
  (e: 'dirty-change', payload: DirtyChangePayload): void
  /** 图片增删移改(前后 ImageAnchor) */
  (e: 'image-change', payload: ImageChangePayload): void
  /** 行列结构变更(增删行列 / 撤销重做的整体还原) */
  (e: 'struct-change', payload: StructChangePayload): void
  /** 权限拒绝(Phase A, 2026-06-08):粘贴/合并/图片落点等命中只读 → 默认 skip + 此事件通知 */
  (e: 'permission-denied', payload: PermissionDeniedPayload): void
  /** 右键菜单触发前(Plan C):用户调 `payload.preventDefault()` 阻止内置菜单(然后自渲染) */
  (e: 'before-context-menu', payload: ContextMenuBeforePayload): void
  /** 右键菜单"展示"通知(Plan C):无论内置是否弹都触发,供自渲染或事件流串到业务 */
  (e: 'context-menu', payload: ContextMenuShowPayload): void
}>()

const { loading, error, workbook, load, loadModel, progress, sourceBuffer } = useExcelDocument()

/** 把 :workbook prop 转成 WorkbookModel(支持 WorkbookModel 直传 + JsonInput 自动构造)。 */
function resolveWorkbookInput(w: WorkbookModel | JsonInput | undefined): WorkbookModel | null {
  if (!w) return null
  return isWorkbookModel(w) ? (w as WorkbookModel) : jsonToWorkbook(w as JsonInput, props.jsonOptions)
}

// 运行时模板状态(P3 进阶):工具栏内置 'template' 项导入的 .xlsx 优先覆盖 :templateFile prop;
// 用户给 `:fileName` / `:templateName` 时仍按 props 走,但 demo / 工具栏导入会自动用 File.name 兜底。
const runtimeTemplateSrc = ref<ExcelSource | null>(null)
const runtimeTemplateName = ref<string | null>(null)
const effectiveTemplateSrc = computed(() => runtimeTemplateSrc.value ?? props.templateFile ?? null)
const effectiveTemplateName = computed(() => runtimeTemplateName.value ?? props.templateName ?? '')
/** demo / 工具栏导入模板用:把文件喂给运行时,重新计算渲染管线 */
function setRuntimeTemplate(src: ExcelSource | null, name: string | null) {
  runtimeTemplateSrc.value = src
  runtimeTemplateName.value = name
}
function clearRuntimeTemplate() {
  setRuntimeTemplate(null, null)
}

// 工具栏内置 'template' 项的隐藏文件拾取器(导入 .xlsx → 设运行时模板)
const templateInputEl = ref<HTMLInputElement | null>(null)
function openTemplateFilePicker() {
  templateInputEl.value?.click()
}
function onTemplateFilePicked(e: Event) {
  const f = (e.target as HTMLInputElement).files?.[0]
  if (!f) return
  setRuntimeTemplate(f, f.name)
  // 清空 input 让同一文件可以再次选(重渲)
  ;(e.target as HTMLInputElement).value = ''
}
/** 显示用文件名:`:fileName` > JSON 源缺省 "JSON 数据" > 首表表名 > '未命名工作簿'(留给 ViewerToolbar 兜底) */
const displayFileName = computed(() => {
  if (props.fileName) return props.fileName
  // :workbook 给了(JSON / 模型) → 默认 "JSON 数据"(用户没传名时)
  if (props.workbook) return 'JSON 数据'
  return workbook.value?.sheets[0]?.name || ''
})

/**
 * 解析一个 .xlsx 源(File / Blob / ArrayBuffer / URL)成 WorkbookModel,**不**改 activeWorkbook —— 用作模板加载。
 */
async function parseTemplateFile(src: ExcelSource): Promise<WorkbookModel> {
  const buffer = await loadArrayBuffer(src)
  const fmt = detectFormat(buffer)
  if (fmt === 'xls') throw new Error('模板文件是旧版 .xls 或加密,仅支持 .xlsx/.xlsm')
  if (fmt === 'not-zip') throw new Error('模板文件不是有效的 .xlsx(非 ZIP 包)')
  if (fmt === 'empty') throw new Error('模板文件为空')
  const model = await parseInWorker(buffer)
  finalizeImages(model)
  return model
}

/**
 * 统一的数据/模板加载入口(P3 重设计 2026-06-08)。
 *   - `:src`(xlsx)→ 直接 load;若同时设了 `:templateFile`,console.warn 并忽略(xlsx 自带格式)
 *   - `:workbook` + 模板 → 解析模板 → applyStyleTemplate 套样式 → loadModel(merged)
 *   - `:workbook` 无模板 → loadModel(纯 JSON,默认样式)
 *   - 只给 `:templateFile`(没数据)→ 把模板当文件加载(用户可能只想预览模板)
 */
async function runInitialLoad() {
  const tplSrc = effectiveTemplateSrc.value
  const initDataWb = resolveWorkbookInput(props.workbook)

  if (props.src) {
    if (tplSrc) console.warn('[ooxml-excel-editor] :templateFile 只在 :workbook (JSON / 模型) 数据源下生效;xlsx 数据源已自带格式,模板已忽略.')
    await load(props.src, effectiveTransform)
    return
  }

  if (initDataWb && tplSrc) {
    try {
      const tplWb = await parseTemplateFile(tplSrc)
      const merged = applyStyleTemplate(initDataWb, tplWb)
      loadModel(merged, effectiveTransform)
    } catch (e) {
      console.error('[ooxml-excel-editor] 模板加载失败,降级为纯 JSON 渲染:', e)
      loadModel(initDataWb, effectiveTransform)
    }
    return
  }

  if (initDataWb) {
    loadModel(initDataWb, effectiveTransform)
    return
  }

  if (tplSrc) await load(tplSrc, effectiveTransform)
}

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
const sheetsVersion = ref(0) // 工作表列表版本(core 增删表时 +1;workbook 为 shallowRef,用它强制 SheetTabs 重渲)
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
const editorSlotEl = ref<HTMLElement | null>(null) // E2: 单元格编辑器挂载层
let controller: ViewerController | null = null

function doRender() {
  controller?.render()
}
function measure() {
  controller?.measure()
}

function rebuildRenderer() {
  const s = sheet.value
  const wb = workbook.value
  if (!s || !wb || !controller) return
  // 控制器内部负责清选区/查找命中/tooltip,并在末尾按需重跑当前查找
  controller.rebuild(s, wb, zoom.value, {
    theme: effectiveTheme.value,
    cellStyle: hasCellStyleHook.value ? combinedCellStyle : undefined,
    cellImageFit: props.cellImageFit,
    readOnlyCellStyle: props.readOnlyCellStyle, // Phase C 2026-06-08
  })
  controller.setSourceBuffer(sourceBuffer.value) // 注入原件字节(overlay 高保真导出)
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
  if (canvasEl.value && renderAreaEl.value && scrollerEl.value && spacerEl.value && editorSlotEl.value && ovMain.value && ovFRow.value && ovFCol.value && ovCorner.value) {
    controller = new ViewerController(
      {
        canvas: canvasEl.value,
        renderArea: renderAreaEl.value,
        scroller: scrollerEl.value,
        spacer: spacerEl.value,
        overlays: { main: ovMain.value, frow: ovFRow.value, fcol: ovFCol.value, corner: ovCorner.value },
        editorSlot: editorSlotEl.value,
      },
      {
        onRenderer: (r) => (renderer.value = r),
        onRenderTick: () => renderTick.value++,
        onSelectionChange: () => selVersion.value++,
        onCellClick: (row, col, text) => fire('cell-click', { row, col, text }),
        onCellDblClick: (row, col, text) => fire('cell-dblclick', { row, col, text }),
        onHyperlink: (url, cell) => {
          fire('hyperlink-click', { url, cell })
          if (props.openLinks) window.open(url, '_blank', 'noopener')
        },
        onTooltip: (tip) => (tooltip.value = tip),
        onFindChange: () => findVersion.value++,
        onFilterChange: () => filterVersion.value++,
        onActiveSheetChange: (index) => { activeSheet.value = index; sheetsVersion.value++ }, // core 新增表(透视新建工作表)→ 顶版本号让 SheetTabs 重读(workbook 是 shallowRef,push 不自动通知)
        onEditEvent: (event, payload) => fire(event, payload),
        onContextMenuBefore: (payload) => {
          // 1) 先跑插件 contextMenu(数组顺序串行,后者拿前者的输出)
          for (const p of normalizedPlugins.value) {
            if (p.contextMenu) {
              const next = p.contextMenu(payload.ctx, payload.items)
              if (Array.isArray(next)) payload.items.splice(0, payload.items.length, ...next)
            }
          }
          // 2) `:context-menu="false"` 直接阻止内置弹层(但事件仍会触发 onContextMenuShow,供用户自渲染)
          if (props.contextMenu === false) payload.preventDefault()
          // 3) emit 事件(在阻止判定后,允许用户 listener 进一步 preventDefault)
          ;(emit as (e: string, p: unknown) => void)('before-context-menu', payload)
        },
        onContextMenuShow: (payload) => (emit as (e: string, p: unknown) => void)('context-menu', payload),
      },
    )
    view.value = controller.view // 壳与控制器共享同一 view 对象(现有 view.value 读法不变)
    controller.fileName = props.fileName // 导出默认文件名
    controller.setEditConfig(effectiveEditConfig.value) // 编辑配置(默认只读)
    controller.setEditorResolver(hasEditor.value ? resolveEditor : undefined) // E2: 编辑器解析
    controller.setLightboxEnabled(props.imageLightbox !== false) // 图片点击放大(默认开)
    // 右键菜单 transform(Plan C):函数形式 = 用户 prop 直接覆盖内置(在插件 transform 之后再跑)
    controller.setContextMenuTransform(typeof props.contextMenu === 'function' ? props.contextMenu : null)
  }
  if (pluginOvEl.value) pluginOverlayHost = new PluginOverlayHost(pluginOvEl.value)
  void runInitialLoad()
  resizeObserver = new ResizeObserver(() => {
    measure()
    doRender()
  })
  if (renderAreaEl.value) resizeObserver.observe(renderAreaEl.value)
})

onBeforeUnmount(() => {
  resizeObserver?.disconnect()
  controller?.dispose()
  pluginOverlayHost?.dispose()
  pluginCleanups.forEach((fn) => fn())
  if (workbook.value) revokeImages(workbook.value)
})

// src / workbook / templateFile / runtimeTemplate 任一变化 → 重跑统一入口
watch([() => props.src, () => props.workbook, () => props.templateFile, runtimeTemplateSrc], () => {
  void runInitialLoad()
})

watch(() => props.fileName, (f) => {
  if (controller) controller.fileName = f
})

watch(effectiveEditConfig, (cfg) => controller?.setEditConfig(cfg))
watch(() => props.contextMenu, (cm) => controller?.setContextMenuTransform(typeof cm === 'function' ? cm : null))
watch([() => props.editor, normalizedPlugins], () => controller?.setEditorResolver(hasEditor.value ? resolveEditor : undefined))
watch(() => props.cellImageFit, (fit) => { if (fit) controller?.setCellImageFit(fit) })
watch(() => props.imageLightbox, (v) => controller?.setLightboxEnabled(v !== false))

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
  controller?.clearFilterState() // 新工作簿: 旧筛选态作废(模型已换)
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
  // 离开旧表: 恢复其被筛选隐藏的行(oldIdx 此刻仍指向旧表模型)
  if (oldIdx != null) controller?.resetFilter(workbook.value?.sheets[oldIdx])
  await nextTick()
  rebuildRenderer()
})

watch(zoom, (z) => controller?.setZoom(z))

// ---------------- 交互: 选区 / 超链接 / 悬停 / 复制(模型与交互在 ViewerController) ----------------
// 选区模型已下沉控制器;壳只持 selVersion(变化时让派生计算属性重算)+ tooltip(控制器经 onTooltip 回填)
const selVersion = ref(0)
const tooltip = ref<TooltipState | null>(null)

const selection = computed<MergeRange | null>(() => {
  void selVersion.value // 选区模型变化 → 重算
  void renderer.value // renderer 就绪后重算
  return controller?.getSelection() ?? null
})

const activeCellAddr = computed(() => {
  void selVersion.value
  const c = controller?.getActiveCell()
  return c ? colIndexToLetters(c.col) + (c.row + 1) : ''
})
const selRangeLabel = computed(() => {
  const s = selection.value
  if (!s || (s.top === s.bottom && s.left === s.right)) return ''
  return `${colIndexToLetters(s.left)}${s.top + 1}:${colIndexToLetters(s.right)}${s.bottom + 1}`
})
const formulaBarText = computed(() => {
  void selVersion.value
  const r = renderer.value
  const c = controller?.getActiveCell()
  if (!r || !c) return ''
  return r.cellFormula(c.row, c.col) ?? r.cellText(c.row, c.col)
})
// 公式栏可编辑 + 与单元格联动(提交→改格;切格/格内编辑→栏更新)
const fbDraft = ref('')
const fbEditing = ref(false)
// Phase 1.2.1 (2026-06-08) 公式栏自动撑高: textarea ref + input/draft 变化时同步高度
const fbEl = ref<HTMLTextAreaElement | null>(null)
function syncFbHeight() {
  const el = fbEl.value
  if (!el) return
  el.style.height = 'auto'
  el.style.height = el.scrollHeight + 'px'
}
function onFbInput(e: Event) {
  fbDraft.value = (e.target as HTMLTextAreaElement).value
  syncFbHeight()
}
const fbCanEdit = computed(() => {
  void selVersion.value
  void renderTick.value
  return !!controller?.canEditActiveCell()
})
const formulaBarEditString = computed(() => {
  void selVersion.value
  void renderTick.value
  return controller?.getCellEditString() ?? ''
})
watch(formulaBarEditString, (v) => { if (!fbEditing.value) fbDraft.value = v }, { immediate: true })
// 切格 / 切表 / 内容外部变更 → fbDraft 变 → 撑高同步 (用 nextTick 保证 textarea 渲染完)
watch(fbDraft, () => nextTick(syncFbHeight))
function fbFocus() {
  fbEditing.value = true
  fbDraft.value = formulaBarEditString.value
  nextTick(syncFbHeight)
}
function fbCommit(move?: 'down') {
  controller?.commitActiveCellValue(fbDraft.value, move)
  fbEditing.value = false
  fbDraft.value = formulaBarEditString.value
  if (move === 'down') scrollerEl.value?.focus()
}
function fbCancel() {
  fbEditing.value = false
  fbDraft.value = formulaBarEditString.value
  scrollerEl.value?.focus()
}
function fbBlur() {
  if (fbEditing.value) fbCommit()
}
function fbKeydown(e: KeyboardEvent) {
  e.stopPropagation() // 别让网格键盘处理插手
  if (e.key === 'Enter') {
    e.preventDefault()
    fbCommit('down')
  } else if (e.key === 'Escape') {
    e.preventDefault()
    fbCancel()
  }
}
const stats = computed(() => {
  void selVersion.value
  const r = renderer.value
  const s = controller?.getSelection() ?? null
  return r && s ? r.selectionStats(s) : null
})

function fmtNum(n: number): string {
  if (!isFinite(n)) return '—'
  return n.toLocaleString('en-US', { maximumFractionDigits: 2 })
}

// 模板事件 → 控制器(薄包装,避免模板里直接绑 controller 的 this/null 问题)
function onMouseDown(e: MouseEvent) {
  controller?.onMouseDown(e)
}
function onContextMenu(e: MouseEvent) {
  controller?.onContextMenu(e)
}
function onMouseMove(e: MouseEvent) {
  controller?.onMouseMove(e)
}
function onMouseUp(e: MouseEvent) {
  controller?.onMouseUp(e)
}
function onMouseLeave() {
  controller?.onMouseLeave()
}
function onDblClick(e: MouseEvent) {
  controller?.onDblClick(e)
}
function onKeyDown(e: KeyboardEvent) {
  controller?.onKeyDown(e)
}

// ---------------- 选区变化事件 + overlay 定位 API ----------------
watch(selection, (sel) => {
  const active = controller?.getActiveCell()
  if (sel && active) fire('selection-change', { range: sel, active })
})

/** 单元格当前屏幕矩形(render-area 相对坐标);供 overlay slot / 命令式定位 */
function rectOf(row: number, col: number): { x: number; y: number; w: number; h: number } | null {
  return controller?.rectOf(row, col) ?? null
}
/** 区域当前屏幕矩形(左上到右下的并集) */
function rectOfRange(range: MergeRange): { x: number; y: number; w: number; h: number } | null {
  return controller?.rectOfRange(range) ?? null
}

// ---------------- 导出 / 打印(编排在 core/export/WorkbookExporter,控制器委托) ----------------
// 进度遮罩状态(P1.5):内置 onProgress + AbortController,与用户传入的 onProgress/signal 链接
const exportState = ref<ExportProgress | null>(null)
const exportBusy = ref(false)
let exportCtrl: AbortController | null = null
function cancelExport() { exportCtrl?.abort() }

/** 包一层:① 建内置 AbortController + 接 onProgress 给 overlay;② 与用户的 opts.onProgress / opts.signal 链接;
 *  ③ 失败/取消/完成都关 overlay。`exportProgress=false` 时直接透传,纯回调走原路。 */
function chain<T, O extends { onProgress?: (p: ExportProgress) => void; signal?: AbortSignal } | undefined>(
  userOpts: O,
  run: (opts: O) => Promise<T>,
): Promise<T> {
  if (props.exportProgress === false) return run(userOpts)
  const ctrl = new AbortController()
  exportCtrl = ctrl
  if (userOpts?.signal) {
    if (userOpts.signal.aborted) ctrl.abort()
    else userOpts.signal.addEventListener('abort', () => ctrl.abort(), { once: true })
  }
  exportBusy.value = true
  exportState.value = null
  const onProgress = (p: ExportProgress) => {
    exportState.value = p
    userOpts?.onProgress?.(p)
  }
  const merged = { ...(userOpts ?? {}), onProgress, signal: ctrl.signal } as O
  return run(merged).finally(() => {
    exportBusy.value = false
    exportState.value = null
    exportCtrl = null
  })
}

const exportImage = (opts?: ImageExportOptions): Promise<Blob> => chain(opts, (o) => controller!.exportImage(o))
const downloadImage = (opts?: ImageExportOptions): Promise<void> => chain(opts, (o) => controller!.downloadImage(o))
const exportPdf = (opts?: PdfExportOptions): Promise<Blob> => chain(opts, (o) => controller!.exportPdf(o))
const downloadPdf = (opts?: PdfExportOptions): Promise<void> => chain(opts, (o) => controller!.downloadPdf(o))
const print = (opts?: PrintOptions): Promise<void> => chain(opts, (o) => controller!.print(o))

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
  console.error('[ooxml-excel-editor] 导出失败:', e)
  emit('error', msg)
  if (typeof window !== 'undefined' && window.alert) window.alert(msg)
}

// ---- 查找(状态/逻辑在 ViewerController) ----
const findOpen = ref(false) // 纯 UI: FindBar 是否展开
const findVersion = ref(0) // 控制器 onFindChange 回调 +1
const findState = computed<FindState>(() => {
  void findVersion.value
  return controller?.getFindState() ?? { query: '', matchCase: false, wholeCell: false, count: 0, index: -1, replace: '' }
})
function openFind() {
  findOpen.value = true
}
function closeFind() {
  findOpen.value = false
  controller?.clearFind()
  scrollerEl.value?.focus()
}
/** 根容器捕获 Ctrl/Cmd+F → 打开查找(替代浏览器原生查找) */
function onRootKeydown(e: KeyboardEvent) {
  if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) {
    e.preventDefault()
    openFind()
  }
}

// ---- 自动筛选(状态/逻辑在 ViewerController) ----
const filterVersion = ref(0) // 控制器 onFilterChange 回调 +1
const filterPopup = computed(() => {
  void filterVersion.value
  return controller?.getFilterPopup() ?? null
})
function toggleAutoFilter() {
  controller?.toggleAutoFilter()
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
    const c = controller?.getActiveCell()
    s.freeze = { frozenRows: c ? c.row : 1, frozenCols: c ? c.col : 0 }
  }
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
  controller?.setSelectionRange(range)
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
  scrollToCell: (row, col, opts) => controller?.scrollToCell(row, col, opts) ?? false,
  rectOf,
  rectOfRange,
  redraw: () => doRender(),
  isCellEditable: (row, col) => controller?.isCellEditable(row, col) ?? false,
  setEditableTargets: (targets) => controller?.setEditableTargets(targets),
  getEditableTargets: () => controller?.getEditableTargets(),
  sortActiveColumn: (dir) => controller?.sortActiveColumn(dir) ?? false,
  createPivotTable: (opts) => controller?.createPivotTable(opts) ?? false,
  createPivotTableFromSelection: (opts) => controller?.createPivotTableFromSelection(opts) ?? false,
  openPivotTableDialog: () => controller?.openPivotTableDialog() ?? false,
  getConditionalRules: () => controller?.getConditionalRules() ?? [],
  addConditionalRule: (rule) => controller?.addConditionalRule(rule) ?? false,
  updateConditionalRule: (ruleId, patch) => controller?.updateConditionalRule(ruleId, patch) ?? false,
  removeConditionalRule: (ruleId) => controller?.removeConditionalRule(ruleId) ?? false,
  setConditionalRules: (rules) => controller?.setConditionalRules(rules) ?? false,
  openConditionalFormatDialog: () => controller?.openConditionalFormatDialog() ?? false,
  setSelectionNumberFormat: (code) => controller?.setSelectionNumberFormat(code) ?? false,
  openNumberFormatDialog: () => controller?.openNumberFormatDialog() ?? false,
  getCellComment: (row, col) => controller?.getCellComment(row, col) ?? '',
  setCellComment: (row, col, comment) => controller?.setCellComment(row, col, comment) ?? false,
  openCommentEditor: (row, col) => controller?.openCommentEditor(row, col) ?? false,
  editCell: (row, col, value) => controller?.editCell(row, col, value) ?? false,
  editRange: (range, values) => controller?.editRange(range, values) ?? false,
  clearRange: (range) => controller?.clearRange(range) ?? false,
  setStyle: (range, patch) => controller?.setStyle(range, patch) ?? false,
  getActiveFillColor: () => controller?.getActiveFillColor() ?? '#FFFFFF',
  getActiveFontColor: () => controller?.getActiveFontColor() ?? '#000000',
  setSelectionFill: (color) => controller?.setSelectionFill(color) ?? false,
  setSelectionFontColor: (color) => controller?.setSelectionFontColor(color) ?? false,
  getSelectionWrapState: () => controller?.getSelectionWrapState() ?? 'none',
  toggleWrapTextOnSelection: () => controller?.toggleWrapTextOnSelection() ?? false,
  mergeCells: (range) => controller?.mergeCells(range) ?? false,
  unmergeCells: (range) => controller?.unmergeCells(range) ?? false,
  pasteText: (text, at) => controller?.pasteText(text, at) ?? false,
  pasteRichHtml: (html, at, behaviorOverride) => controller?.pasteRichHtml(html, at, behaviorOverride) ?? false,
  getPasteBehavior: () => controller?.getPasteBehavior() ?? DEFAULT_PASTE_BEHAVIOR,
  setPasteBehavior: (cfg) => controller?.setPasteBehavior(cfg),
  openPasteConfigDialog: () => controller?.openPasteConfigDialog() ?? false,
  pasteImageBlob: (blob, at) => controller?.pasteImageBlob(blob, at) ?? Promise.resolve(false),
  getImages: () => controller?.getImages() ?? [],
  addImage: (a) => controller?.addImage(a) ?? -1,
  removeImage: (i) => controller?.removeImage(i) ?? false,
  moveImage: (i, dx, dy) => controller?.moveImage(i, dx, dy) ?? false,
  resizeImage: (i, w, h) => controller?.resizeImage(i, w, h) ?? false,
  getCellEditString: () => controller?.getCellEditString() ?? '',
  canEditActiveCell: () => controller?.canEditActiveCell() ?? false,
  commitActiveCellValue: (value, move) => controller?.commitActiveCellValue(value, move) ?? false,
  getCellImages: () => controller?.getCellImages() ?? [],
  getCellImageAt: (row, col) => controller?.getCellImageAt(row, col) ?? null,
  openImageLightbox: (src, fileName, mime) => controller?.openImageLightbox(src, fileName, mime),
  setCellImageFit: (fit) => controller?.setCellImageFit(fit),
  convertImageToCell: (i, row, col) => controller?.convertImageToCell(i, row, col) ?? false,
  convertImageToCellAuto: (i) => controller?.convertImageToCellAuto(i) ?? false,
  convertAllImagesToCells: (col) => controller?.convertAllImagesToCells(col) ?? 0,
  convertImagesInRangeToCell: (range) =>
    chain<number, { onProgress?: (p: ExportProgress) => void; signal?: AbortSignal }>({}, async (o) => {
      o.onProgress?.({ stage: 'convert', label: '选区浮动图批量嵌入…', ratio: undefined })
      return controller?.convertImagesInRangeToCell(range) ?? 0
    }),
  convertCellImagesInRangeToFloat: (range, size) =>
    chain<number, { onProgress?: (p: ExportProgress) => void; signal?: AbortSignal }>({}, async (o) => {
      o.onProgress?.({ stage: 'convert', label: '选区内嵌图批量浮动化…', ratio: undefined })
      return controller?.convertCellImagesInRangeToFloat(range, size) ?? 0
    }),
  openContextMenu: (x, y, items) => controller?.openContextMenu(x, y, items),
  closeContextMenu: () => controller?.closeContextMenu(),
  convertCellImageToFloat: (row, col, size) => controller?.convertCellImageToFloat(row, col, size) ?? false,
  insertRows: (at, count) => controller?.insertRows(at, count) ?? false,
  deleteRows: (at, count) => controller?.deleteRows(at, count) ?? false,
  insertCols: (at, count) => controller?.insertCols(at, count) ?? false,
  deleteCols: (at, count) => controller?.deleteCols(at, count) ?? false,
  undo: () => controller?.undo(),
  redo: () => controller?.redo(),
  canUndo: () => controller?.canUndo() ?? false,
  canRedo: () => controller?.canRedo() ?? false,
  getEditingCell: () => controller?.getEditingCell() ?? null,
  getCellSnapshot: (row, col) => controller?.getCellSnapshot(row, col) ?? null,
  inspectCell: (row, col) => controller?.inspectCell(row, col) ?? null,
  beginEdit: (row, col) => controller?.beginEdit(row, col) ?? false,
  cancelEdit: () => controller?.cancelEdit(),
  isEditing: () => controller?.isEditing() ?? false,
  setColumnWidth: (target, width) => controller?.setColumnWidth(target, width) ?? 0,
  setRowHeight: (target, height) => controller?.setRowHeight(target, height) ?? 0,
  autoFitColumns: (target) => controller?.autoFitColumns(target) ?? 0,
  autoFitRows: (target) => controller?.autoFitRows(target) ?? 0,
  resetColumnWidth: (target) => controller?.resetColumnWidth(target) ?? 0,
  resetRowHeight: (target) => controller?.resetRowHeight(target) ?? 0,
  isRecalcReady: () => controller?.isRecalcReady() ?? false,
  getVirtualExtent: () => controller?.getVirtualExtent() ?? { rows: 0, cols: 0 },
  isDirty: () => controller?.isDirty() ?? false,
  resetToOriginal: () => controller?.resetToOriginal() ?? false,
  exportImage,
  downloadImage,
  exportPdf,
  downloadPdf,
  print,
  exportXlsx: (opts) => chain(opts, (o) => controller!.exportXlsx(o)),
  downloadXlsx: (opts) => chain(opts, (o) => controller!.downloadXlsx(o)),
  exportJson: (opts) => controller?.exportJson(opts) ?? '{}',
  downloadJson: (opts) => controller?.downloadJson(opts),
  exportCsv: (opts) => controller?.exportCsv(opts) ?? '',
  downloadCsv: (opts) => controller?.downloadCsv(opts),
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
        disabled: !controller?.hasFilters(),
        onClick: () => controller?.clearAllFilters(),
      })
    case 'sort': {
      const sortState = controller?.getSortState()
      const active = controller?.getActiveCell()
      const disabled = !active || !sheet.value
      return bi({
        id,
        iconSvg: I('sort'),
        label: '排序',
        title: active ? `按 ${colIndexToLetters(active.col)} 列排序` : '选中一个单元格后按该列排序',
        active: !!(active && sortState?.col === active.col && sortState.dir),
        disabled,
        items: [
          bi({
            id: 'sort-asc',
            label: '升序 (A → Z / 小 → 大)',
            active: !!(active && sortState?.col === active.col && sortState.dir === 'asc'),
            disabled,
            onClick: () => controller?.sortActiveColumn('asc'),
          }),
          bi({
            id: 'sort-desc',
            label: '降序 (Z → A / 大 → 小)',
            active: !!(active && sortState?.col === active.col && sortState.dir === 'desc'),
            disabled,
            onClick: () => controller?.sortActiveColumn('desc'),
          }),
        ],
      })
    }
    case 'copy':
      return bi({
        id,
        iconSvg: I('copy'),
        label: '复制',
        title: '复制选区 (Ctrl+C)',
        disabled: !selection.value,
        onClick: () => void controller?.copySelection(),
      })
    case 'pivot-table':
      if (!props.pivotTable) return null // 功能未开启(默认):不渲染入口
      return bi({
        id,
        iconSvg: I('pivot-table'),
        label: '透视表',
        title: '选择字段并基于当前选区创建静态透视汇总表',
        disabled: !selection.value || !props.editable,
        onClick: () => controller?.openPivotTableDialog(),
      })
    case 'conditional-format':
      if (!props.conditionalFormat) return null // 功能未开启(默认):不渲染入口
      return bi({
        id,
        iconSvg: I('conditional-format'),
        label: '条件格式',
        title: '管理条件格式规则(新建/编辑/删除;新建套到当前选区)',
        disabled: !props.editable,
        onClick: () => controller?.openConditionalFormatDialog(),
      })
    case 'number-format':
      return bi({
        id,
        iconSvg: I('number-format'),
        label: '数字格式',
        title: '设置单元格数字格式(数值/货币/百分比/日期/自定义)',
        disabled: !selection.value || !props.editable,
        onClick: () => controller?.openNumberFormatDialog(),
      })
    case 'wrap-text': {
      const wrapState = controller?.getSelectionWrapState() ?? 'none'
      return bi({
        id,
        iconSvg: I('wrap-text'),
        label: '自动换行',
        title: '自动换行(选区,WPS 风格 toggle)',
        active: wrapState === 'all',
        disabled: !selection.value || !props.editable,
        onClick: () => void controller?.toggleWrapTextOnSelection(),
      })
    }
    case 'template': {
      const active = !!effectiveTemplateSrc.value
      const name = effectiveTemplateName.value
      // 模板只在 JSON / 模型数据源下生效;xlsx 数据源(`:src`)走自己的格式,套模板无意义,禁用 UI
      const isXlsxSrc = !!props.src && !props.workbook
      return bi({
        id,
        iconSvg: I('template'),
        label: '模板',
        title: isXlsxSrc
          ? '模板仅对 JSON / 模型数据源生效;当前是 .xlsx 数据源,模板不可用'
          : active
            ? `模板已加载:${name || '(未命名)'}`
            : '为 JSON / 模型数据源套用 .xlsx 模板的样式(边框 / 字体 / 列宽 / 合并 等);模板的文字内容会被丢弃',
        active,
        disabled: isXlsxSrc,
        items: [
          bi({
            id: 'tpl-default',
            label: (!active ? '✓ ' : '') + '默认渲染',
            title: '不套模板,数据按默认样式渲染',
            disabled: !active,
            onClick: clearRuntimeTemplate,
          }),
          bi({ id: 'tpl-sep', type: 'separator' }),
          bi({
            id: 'tpl-import',
            label: '导入 .xlsx 模板…',
            title: '选一份 .xlsx,把它的 styling(边框/字体/列宽/合并/freeze) 套到当前 JSON 数据上;模板的文字内容会被丢弃',
            onClick: openTemplateFilePicker,
          }),
          bi({
            id: 'tpl-clear',
            label: '清除模板',
            title: '切回默认样式渲染',
            disabled: !active,
            onClick: clearRuntimeTemplate,
          }),
        ],
      })
    }
    case 'image-tools': {
      const sel = selection.value
      const active = controller?.getActiveCell()
      const hasFloats = (sheet.value?.images.length ?? 0) > 0
      return bi({
        id,
        iconSvg: I('image-tools'),
        label: '图片工具',
        title: '浮动图 ⇄ 单元格内嵌图(WPS DISPIMG)互转',
        disabled: !props.editable,
        items: [
          bi({
            id: 'img-sel-to-cell',
            label: '选区:浮动 → 嵌入',
            title: '把选区里"中心格在选区内"的浮动图,就近嵌入',
            disabled: !sel || !hasFloats,
            onClick: () => sel && controller?.convertImagesInRangeToCell(sel),
          }),
          bi({
            id: 'img-sel-to-float',
            label: '选区:嵌入 → 浮动',
            title: '把选区内所有 DISPIMG 格拎成浮动图',
            disabled: !sel,
            onClick: () => sel && controller?.convertCellImagesInRangeToFloat(sel),
          }),
          bi({ id: 'img-sep', type: 'separator' }),
          bi({
            id: 'img-all-to-cell',
            label: '整表:浮动 → 嵌入',
            title: '全表浮动图按几何就近嵌入各自单元格',
            disabled: !hasFloats,
            onClick: () => controller?.convertAllImagesToCells(),
          }),
          bi({
            id: 'img-col-to-cell',
            label: '整列:浮动 → 嵌入(活动列)',
            title: '把中心落在活动列的浮动图就近嵌入',
            disabled: !hasFloats || !active,
            onClick: () => active && controller?.convertAllImagesToCells(active.col),
          }),
        ],
      })
    }
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
      return null
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
  // 选区/查找/筛选/渲染状态变化时重算 active/disabled/label
  void renderTick.value
  void selVersion.value
  void findVersion.value
  void filterVersion.value
  if (props.toolbar === false) return []
  const entries: Array<string | ToolbarItem> = Array.isArray(props.toolbar) ? props.toolbar : ['find', 'filter', 'sort']
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

// 插件 overlay 用框架无关的 PluginOverlayHost 挂 DOM(随 renderTick 重渲),与 React 壳共用
const pluginOvEl = ref<HTMLElement | null>(null)
let pluginOverlayHost: PluginOverlayHost | null = null
function renderPluginOverlays() {
  const ctx: OverlayContext = { rectOf, rectOfRange, tick: renderTick.value, workbook: workbook.value }
  pluginOverlayHost?.render(normalizedPlugins.value, ctx)
}
watch([renderTick, normalizedPlugins], renderPluginOverlays, { flush: 'post' })
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
        :file-name="displayFileName"
        :template-name="effectiveTemplateName"
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
      <!-- Phase 1.2.1 (2026-06-08): input → textarea, 长公式 / 多行内容自动撑高 (WPS 风格).
           Shift+Enter 插换行, 普通 Enter 提交;上限 ~6 行 (max-height CSS 控). -->
      <textarea
        v-if="fbCanEdit"
        ref="fbEl"
        class="content content-input"
        :value="fbDraft"
        :title="fbDraft"
        rows="1"
        spellcheck="false"
        @focus="fbFocus"
        @input="onFbInput"
        @keydown="fbKeydown"
        @blur="fbBlur"
      />
      <span v-else class="content" :title="formulaBarText">{{ formulaBarText }}</span>
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
        @contextmenu="onContextMenu"
      >
        <div class="spacer" ref="spacerEl" />
      </div>

      <FindBar
        v-if="findOpen && workbook"
        :query="findState.query"
        :match-count="findState.count"
        :current="findState.index"
        :match-case="findState.matchCase"
        :whole-cell="findState.wholeCell"
        :editable="!!props.editable"
        :replace="findState.replace"
        @update:query="controller?.setFindQuery($event)"
        @update:match-case="controller?.setFindMatchCase($event)"
        @update:whole-cell="controller?.setFindWholeCell($event)"
        @update:replace="controller?.setFindReplace($event)"
        @next="controller?.findNext()"
        @prev="controller?.findPrev()"
        @replace-one="controller?.replaceCurrent()"
        @replace-all="controller?.replaceAll()"
        @close="closeFind"
      />

      <FilterPopup
        v-if="filterPopup"
        :values="filterPopup.values"
        :selected="filterPopup.selected"
        :x="filterPopup.x"
        :y="filterPopup.y"
        :sort-dir="filterPopup.sortDir"
        @apply="controller?.applyFilterSelection($event)"
        @clear="controller?.clearFilterColumn()"
        @close="controller?.closeFilterPopup()"
        @sort="(dir) => { const c = filterPopup?.col; controller?.closeFilterPopup(); if (c != null) controller?.sortColumn(c, dir) }"
      />

      <!-- 分层 UI: 消费方在格子上叠自己的组件,用 rectOf 定位、tick 触发跟随 -->
      <div class="ov-slot">
        <slot name="overlay" :rect-of="rectOf" :rect-of-range="rectOfRange" :tick="renderTick" />
        <!-- 插件 overlay:框架无关 DOM,由 PluginOverlayHost 挂载 -->
        <div ref="pluginOvEl" />
      </div>
      <!-- 单元格编辑器层(E2):在格 + overlay 之上,CellEditorHost 挂载 -->
      <div class="editor-slot" ref="editorSlotEl" />

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
      :key="sheetsVersion"
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

    <!-- 工具栏「模板」项的隐藏文件拾取器(P3 进阶) -->
    <input ref="templateInputEl" type="file" accept=".xlsx,.xlsm" hidden @change="onTemplateFilePicked" />

    <!-- 内置导出进度遮罩(P1.5):props.exportProgress=false 关闭;插槽 #export-progress 完全自渲染 -->
    <template v-if="exportProgress !== false">
      <slot name="export-progress" :state="exportState" :busy="exportBusy" :cancel="cancelExport">
        <ExportProgressOverlay :state="exportState" :busy="exportBusy" @cancel="cancelExport" />
      </slot>
    </template>
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
/* 单元格编辑器层(E2): 在最上,容器穿透,编辑器本身接收交互.
 * Phase 1 长文本撑高 (2026-06-08): overflow:hidden → visible — 让编辑器能向下溢出原格
 * (跟 WPS 一致), z-index:6 仍最上层, 不影响下方网格 / 冻结窗格 / 滚动条交互. */
.editor-slot {
  position: absolute;
  inset: 0;
  z-index: 6;
  overflow: visible;
  pointer-events: none;
}
.editor-slot :deep(*) {
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
/* 公式栏: Phase 1.2.1 (2026-06-08) 自动撑高跟 textarea 内容一致, 上限 ~6 行 */
.formula-bar {
  display: flex;
  align-items: stretch;
  min-height: 28px;
  flex: 0 0 auto;
  border-bottom: 1px solid #e2e4e7;
  background: #fff;
  font-size: 13px;
}
.formula-bar .addr {
  width: 72px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-right: 1px solid #e2e4e7;
  color: #444;
  font-weight: 600;
  flex: 0 0 auto;
}
.formula-bar .fx {
  width: 34px;
  display: flex;
  align-items: center;
  justify-content: center;
  color: #999;
  font-style: italic;
  flex: 0 0 auto;
}
.formula-bar .content {
  flex: 1;
  padding: 6px 8px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  color: #222;
  font-family: Consolas, 'Courier New', monospace;
  display: flex;
  align-items: center;
  line-height: 1.4;
}
.formula-bar textarea.content {
  /* textarea 在公式栏内: 自动撑高, 最多 ~6 行 (~108px) 后内部滚 */
  white-space: pre-wrap;
  overflow: auto;
  text-overflow: clip;
  resize: none;
  border: none;
  outline: none;
  font-family: Consolas, 'Courier New', monospace;
  font-size: 13px;
  display: block;
  min-height: 22px;
  max-height: 108px;
  align-items: stretch;
  line-height: 1.4;
}
.formula-bar .content-input {
  border: none;
  outline: none;
  height: 100%;
  background: transparent;
  font-size: 13px;
}
.formula-bar .content-input:focus {
  background: #f5fbf7;
  box-shadow: inset 0 0 0 1px #21a366;
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
:deep(.ooxml-pivot-button) {
  position: absolute;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 4px;
  min-width: 0;
  padding: 0 3px 0 6px;
  border: 1px solid #9aa7b2;
  border-radius: 2px;
  background: linear-gradient(#ffffff, #e8edf2);
  color: #1f2329;
  font: 12px/1.2 -apple-system, 'Segoe UI', 'Microsoft YaHei', sans-serif;
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.75);
  pointer-events: auto;
  overflow: hidden;
  cursor: default;
}
:deep(.ooxml-pivot-button:hover) {
  border-color: #6e879e;
  background: linear-gradient(#ffffff, #dde8f4);
}
:deep(.ooxml-pivot-button .ooxml-pivot-label) {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
:deep(.ooxml-pivot-button .ooxml-pivot-caret) {
  flex: 0 0 auto;
  color: #3c4a57;
  font-size: 10px;
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
