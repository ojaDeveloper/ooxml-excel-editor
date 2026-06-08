/**
 * Vue 2 壳 (Vue 2.7+ Composition API + render function) — 跟 Vue 3 / React 壳同构, 共用同一份 core.
 *
 * 为什么 .ts 而非 .vue?
 *   - 项目同时存在 vue@3 (Vue 3 壳的 peer) + vue@2 (alias `vue2`) → @vitejs/plugin-vue2 解析
 *     @vue/compiler-sfc 时拿到 vue@3 的版本, SFC 编译失败. 改用 render function 完全绕开
 *     SFC 编译路径. 工程上跟 React 壳 (.tsx + hook) 几乎同构.
 *
 * 为什么 canvas / overlays / scroller 用 createElement 手动挂载 (而不是 h() 渲染)?
 *   - Vue 2 render function 重渲时 patch 会破坏 controller 持有的 DOM 引用 (函数 ref 先解绑
 *     再重绑, controller.els.canvas 变成 stale 引用挂在已 detach 的旧 canvas 上, paint 时
 *     设的 width/height 落到死节点上, 用户看到的新 canvas 永远是 300x150).
 *   - 修复: 把 controller 管理的 DOM 在 onMounted 时手动 createElement + appendChild 到
 *     .ov-render-area, Vue 完全不碰这些节点. render function 只渲染 chrome (toolbar /
 *     formula bar / sheet tabs / state overlay / findbar) — 这些是纯响应式 UI, 走 Vue patch
 *     没问题. 这跟 React 壳的"imperative DOM" 模式一致, 也契合 controller 框架无关的设计.
 */
// 从 '@vue/composition-api' 拿 Composition API → 同时支持 Vue 2.6.x + 2.7+:
//   - Vue 2.6.x:  消费方装 @vue/composition-api 并 Vue.use(VueCompositionAPI)
//   - Vue 2.7+:   @vue/composition-api 装上后自动 re-export Vue 2.7 内置 API (plugin install 为 noop)
// dev 时 vite alias 把 '@vue/composition-api' 重定向到 vue2 (vue@2.7 dist), 拿 2.7 内置 Composition API.
// build 时 rollup external @vue/composition-api, 让消费者自己解析.
import {
  defineComponent,
  h,
  ref,
  reactive,
  computed,
  watch,
  onMounted,
  onBeforeUnmount,
  shallowRef,
  nextTick,
  getCurrentInstance,
  type PropType,
  type VNode,
} from '@vue/composition-api'
import type { ExportConfig } from '@/components/export-types'
import type { ResolvedToolbarItem } from '@/components/toolbar-types'
import { TOOLBAR_ICONS, svgWrap } from '@/components/toolbar-icons'
import type {
  ContextMenuTransform,
  ExcelPlugin,
  ExcelPluginContext,
  OverlayContext,
  PluginEvent,
  ToolbarItem,
  ViewerApi,
} from '@/core/plugin'
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
import type { EditableTarget, EditConfig } from '@/core/edit/types'
import type { FormulaEngineFactory } from '@/core/formula/engine'
import type { EditorResolver, CellEditorFactory } from '@/core/edit/editor-context'
import { revokeImages } from '@/core/finalize'
import type { ImageExportOptions, PdfExportOptions, PrintOptions } from '@/core/export'
import type { ExportProgress } from '@/core/progress'
import { ViewerController, type FindState, type TooltipState } from '@/core/viewer/controller'
import { colIndexToLetters } from '@/core/layout/grid-metrics'
import { getCellValue, getCellText, getSheetData, getRangeData, sheetToJSON, type ReadOptions } from '@/core/model/data-access'
import { useExcelDocumentVue2 } from './use-excel-document'
import './excel-viewer.css'

/** 字符串 ref 工厂 — render 时挂 string ref (`{ ref: slot.refName }`), Vue 把 DOM 放进 vm.$refs.
 *  为什么不用 function/callback ref?
 *    回调 ref 是 Vue 3 引入、Vue 2.7 backport. Vue 2.6 的 vnode ref 只认字符串 ——
 *    传函数根本不会被调用, slot.value 永远 null, renderArea / fb / templateInput 全拿不到 DOM.
 *  字符串 ref + vm.$refs 在 Vue 2.6 / 2.7 / Vue 3 三个版本都受支持, 是最稳的跨版本路径.
 *  注: vm.$refs 是 live lookup, slot.value 是 getter, 每次读时去 vm.$refs 取最新值. */
function makeDomSlotFactory(vm: { $refs: Record<string, unknown> }) {
  return function domSlot<T extends HTMLElement>(refName: string) {
    return {
      refName,
      get value(): T | null { return (vm.$refs[refName] as T | undefined) ?? null },
    }
  }
}

function ce<K extends keyof HTMLElementTagNameMap>(tag: K, className?: string): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag)
  if (className) el.className = className
  return el
}

export default defineComponent({
  name: 'OoxmlExcelViewer',
  props: {
    src: { type: [String, Object, Blob, ArrayBuffer, Uint8Array] as PropType<ExcelSource>, default: undefined },
    workbook: { type: [Array, Object] as PropType<WorkbookModel | JsonInput>, default: undefined },
    jsonOptions: { type: Object as PropType<JsonLoadOptions>, default: undefined },
    templateFile: { type: [String, Object, Blob, ArrayBuffer, Uint8Array] as PropType<ExcelSource>, default: undefined },
    templateName: { type: String, default: '' },
    fileName: { type: String, default: '' },
    theme: { type: Object as PropType<Record<string, unknown>>, default: undefined },
    transformModel: { type: Function as PropType<TransformModelFn>, default: undefined },
    cellStyle: { type: Function as PropType<CellStyleFn>, default: undefined },
    cellImageFit: { type: String as PropType<'fill' | 'contain' | 'cover'>, default: undefined },
    imageLightbox: { type: Boolean, default: true },
    openLinks: { type: Boolean, default: true },
    plugins: { type: Array as PropType<ExcelPlugin[]>, default: () => [] },
    toolbar: { type: [Boolean, Array] as PropType<boolean | Array<string | ToolbarItem>>, default: true },
    editable: { type: Boolean, default: false },
    cellReadOnly: { type: Function as PropType<(cell: CellModel | null, pos: { row: number; col: number }) => boolean | void>, default: undefined },
    readOnlyRanges: { type: Array as PropType<MergeRange[]>, default: undefined },
    editableTargets: { type: [Array, Object] as PropType<EditableTarget | EditableTarget[]>, default: undefined },
    strictDimensions: { type: Boolean, default: false },
    readOnlyCellStyle: { type: [Boolean, Object, Function] as PropType<boolean | CellStyleOverride | CellStyleFn>, default: false },
    editor: { type: Function as PropType<EditorResolver>, default: undefined },
    recalc: { type: Boolean, default: false },
    formulaEngine: { type: Function as PropType<FormulaEngineFactory>, default: undefined },
    contextMenu: { type: [Boolean, Function] as PropType<boolean | ContextMenuTransform>, default: undefined },
    exportProgress: { type: Boolean, default: true },
  },
  setup(props, { emit, expose, slots }) {
    // ---- 由 Vue 管理的 DOM (chrome + 一个空的 render-area 外壳) ----
    // 三个 slot 走 string ref (Vue 2.6/2.7/Vue 3 都支持), getter 实时读 vm.$refs.
    const vm = getCurrentInstance()?.proxy as unknown as { $refs: Record<string, unknown> }
    const domSlot = makeDomSlotFactory(vm)
    const renderAreaSlot = domSlot<HTMLDivElement>('renderArea')
    const fbSlot = domSlot<HTMLTextAreaElement>('fb')
    const templateInputSlot = domSlot<HTMLInputElement>('templateInput')

    // ---- 由 controller 持有的 DOM (onMounted 时 createElement + appendChild, Vue 完全不碰) ----
    let canvasEl: HTMLCanvasElement | null = null
    let scrollerEl: HTMLDivElement | null = null
    let spacerEl: HTMLDivElement | null = null
    let editorSlotEl: HTMLDivElement | null = null
    let ovMainEl: HTMLDivElement | null = null
    let ovFRowEl: HTMLDivElement | null = null
    let ovFColEl: HTMLDivElement | null = null
    let ovCornerEl: HTMLDivElement | null = null
    let pluginOvEl: HTMLDivElement | null = null

    // ---- 数据加载 + workbook 状态 ----
    const { loading, error, workbook, load, loadModel, progress, sourceBuffer } = useExcelDocumentVue2()
    const activeSheet = ref(0)
    const zoom = ref(1)
    const renderTick = ref(0)
    const selVersion = ref(0)
    const findVersion = ref(0)
    const filterVersion = ref(0)
    const tooltip = ref<TooltipState | null>(null)
    const exportState = ref<ExportProgress | null>(null)
    const exportBusy = ref(false)
    let exportCtrl: AbortController | null = null
    const exportDialogOpen = ref(false)
    // 下拉菜单全局唯一 open id (header 导出菜单 / action toolbar 各项的子菜单 / 「更多」折叠菜单 共用)
    const menuOpenId = ref<string | null>(null)
    function setMenuOpen(id: string | null) { menuOpenId.value = id }
    function toggleMenu(id: string) { menuOpenId.value = menuOpenId.value === id ? null : id }
    function onDocClick(e: MouseEvent) {
      if (!(e.target as HTMLElement)?.closest('[data-tb-menu]')) menuOpenId.value = null
    }
    const exportForm = reactive<ExportConfig>({
      action: 'png',
      scope: 'sheet',
      scale: 2,
      includeHeaders: false,
      gridlines: true,
      format: 'auto',
      orientation: 'auto',
      fitToWidth: true,
      pdfVector: false,
    })

    // ---- 模板 + JSON 数据源 ----
    const runtimeTemplateSrc = ref<ExcelSource | null>(null)
    const runtimeTemplateName = ref<string | null>(null)
    const effectiveTemplateSrc = computed(() => runtimeTemplateSrc.value ?? props.templateFile ?? null)
    const effectiveTemplateName = computed(() => runtimeTemplateName.value ?? props.templateName ?? '')
    const displayFileName = computed(() => {
      if (props.fileName) return props.fileName
      if (props.workbook) return 'JSON 数据'
      return workbook.value?.sheets[0]?.name || ''
    })
    function resolveWorkbookInput(w: WorkbookModel | JsonInput | undefined): WorkbookModel | null {
      if (!w) return null
      return isWorkbookModel(w) ? (w as WorkbookModel) : jsonToWorkbook(w as JsonInput, props.jsonOptions)
    }
    async function parseTemplateFile(src: ExcelSource): Promise<WorkbookModel> {
      const buffer = await loadArrayBuffer(src)
      const fmt = detectFormat(buffer)
      if (fmt === 'xls') throw new Error('模板文件是旧版 .xls 或加密, 仅支持 .xlsx/.xlsm')
      if (fmt === 'not-zip') throw new Error('模板文件不是有效的 .xlsx(非 ZIP 包)')
      if (fmt === 'empty') throw new Error('模板文件为空')
      const model = await parseInWorker(buffer)
      finalizeImages(model)
      return model
    }

    // ---- progress 派生 ----
    const progressLabel = computed(() => {
      const p = progress.value
      if (!p) return '加载中…'
      switch (p.stage) {
        case 'read': return '读取文件'
        case 'unzip': return '解压'
        case 'parse': return '解析'
        case 'finalize': return '处理图片'
        default: return '加载中…'
      }
    })
    const progressPct = computed(() => {
      const p = progress.value
      if (!p || p.ratio == null) return null
      return Math.max(0, Math.min(100, Math.round(p.ratio * 100)))
    })

    // ---- 插件 + 钩子合并 ----
    const normalizedPlugins = computed<ExcelPlugin[]>(() => props.plugins ?? [])
    const effectiveTheme = computed(() => Object.assign({}, ...normalizedPlugins.value.map((p) => p.theme || {}), props.theme || {}))
    const hasCellStyleHook = computed(() => !!props.cellStyle || normalizedPlugins.value.some((p) => p.cellStyle))
    function combinedCellStyle(cell: CellModel, pos: { row: number; col: number }, ctx?: { editable: boolean }) {
      let acc: Record<string, unknown> | undefined
      const apply = (fn?: CellStyleFn) => {
        const o = fn?.(cell, pos, ctx)
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
    const effectiveEditConfig = computed<EditConfig>(() => ({
      editable: props.editable,
      cellReadOnly: props.cellReadOnly,
      readOnlyRanges: props.readOnlyRanges,
      editableTargets: props.editableTargets,
      strictDimensions: props.strictDimensions,
      recalc: props.recalc,
      formulaEngine: props.formulaEngine,
    }))
    function resolveEditor(cell: CellModel | null, pos: { row: number; col: number }): CellEditorFactory | void {
      const fromProp = props.editor?.(cell, pos)
      if (fromProp) return fromProp
      for (const p of normalizedPlugins.value) {
        const f = p.editor?.(cell, pos)
        if (f) return f
      }
    }
    const hasEditor = computed(() => !!props.editor || normalizedPlugins.value.some((p) => p.editor))

    // ---- ViewerController + 渲染管线 ----
    const controllerRef = shallowRef<ViewerController | null>(null)
    let resizeObserver: ResizeObserver | null = null
    let pluginOverlayHost: PluginOverlayHost | null = null
    const pluginHandlers = new Map<PluginEvent, Set<(p: any) => void>>()
    let pluginCleanups: Array<() => void> = []

    function fire(event: PluginEvent, payload: any) {
      ;(emit as any)(event, payload)
      pluginHandlers.get(event)?.forEach((h) => h(payload))
    }

    async function runInitialLoad() {
      const tplSrc = effectiveTemplateSrc.value
      const initDataWb = resolveWorkbookInput(props.workbook)
      if (props.src) {
        if (tplSrc) console.warn('[ooxml-excel-editor/vue2] :templateFile 只在 :workbook 数据源下生效;xlsx 数据源已忽略.')
        await load(props.src, effectiveTransform)
        return
      }
      if (initDataWb && tplSrc) {
        try {
          const tplWb = await parseTemplateFile(tplSrc)
          loadModel(applyStyleTemplate(initDataWb, tplWb), effectiveTransform)
        } catch (e) {
          console.error('[ooxml-excel-editor/vue2] 模板加载失败, 降级纯 JSON:', e)
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

    function rebuildRenderer() {
      const wb = workbook.value
      const controller = controllerRef.value
      if (!wb || !controller) return
      const sheet = wb.sheets[activeSheet.value] ?? wb.sheets[0]
      if (!sheet) return
      controller.rebuild(sheet, wb, zoom.value, {
        theme: effectiveTheme.value as never,
        cellStyle: hasCellStyleHook.value ? combinedCellStyle : undefined,
        cellImageFit: props.cellImageFit,
        readOnlyCellStyle: props.readOnlyCellStyle as never,
      })
      controller.setSourceBuffer(sourceBuffer.value)
    }

    function initPlugins() {
      pluginCleanups.forEach((fn) => fn())
      pluginCleanups = []
      pluginHandlers.clear()
      const register = (event: PluginEvent, fn: (p: any) => void) => {
        let set = pluginHandlers.get(event)
        if (!set) pluginHandlers.set(event, (set = new Set()))
        set.add(fn)
      }
      const ctx: ExcelPluginContext = { viewer: viewerApi as ViewerApi, on: register, redraw: () => controllerRef.value?.render() }
      for (const p of normalizedPlugins.value) {
        if (p.events) for (const [ev, fn] of Object.entries(p.events)) if (fn) register(ev as PluginEvent, fn)
        const cleanup = p.setup?.(ctx)
        if (typeof cleanup === 'function') pluginCleanups.push(cleanup)
      }
    }

    function renderPluginOverlays() {
      const controller = controllerRef.value
      if (!pluginOverlayHost || !controller) return
      const ctx: OverlayContext = {
        rectOf: (r, c) => controller.rectOf(r, c),
        rectOfRange: (range) => controller.rectOfRange(range),
        tick: renderTick.value,
        workbook: workbook.value,
      }
      pluginOverlayHost.render(normalizedPlugins.value, ctx)
    }

    onMounted(() => {
      // ⚠ 跨 Vue 2.6 / 2.7 / Vue 3 兼容: slot.value 是 vm.$refs[name] 的 getter (见 makeDomSlotFactory).
      // Vue 2.6 string ref 走同步赋值, onMounted 一进来就能读. 仍 nextTick 是为统一三版本时序 (Vue 3
      // ref 在 mounted 后再 flush, $refs 异步, 这里 nextTick 兜底保证三版本都拿到).
      nextTick(() => {
        const renderArea = renderAreaSlot.value
        if (!renderArea) {
          console.error('[ooxml-excel-editor/vue2] onMounted: renderArea DOM 没拿到 — vm.$refs.renderArea 仍为空, 这是 bug')
          return
        }

      // 手动创建 controller 管理的所有 DOM. Vue 完全不知道它们 → 不会因为重渲销毁/strip.
      canvasEl = ce('canvas', 'ov-grid-canvas')
      ovMainEl = ce('div', 'ov ov-main')
      ovFColEl = ce('div', 'ov ov-fcol')
      ovFRowEl = ce('div', 'ov ov-frow')
      ovCornerEl = ce('div', 'ov ov-corner')
      const ovPluginWrap = ce('div', 'ov ov-plugin')
      pluginOvEl = ce('div')
      ovPluginWrap.appendChild(pluginOvEl)
      scrollerEl = ce('div', 'ov-scroller')
      scrollerEl.tabIndex = 0
      spacerEl = ce('div', 'ov-spacer')
      scrollerEl.appendChild(spacerEl)
      editorSlotEl = ce('div', 'ov-editor-slot')

      renderArea.appendChild(canvasEl)
      renderArea.appendChild(ovMainEl)
      renderArea.appendChild(ovFColEl)
      renderArea.appendChild(ovFRowEl)
      renderArea.appendChild(ovCornerEl)
      renderArea.appendChild(ovPluginWrap)
      renderArea.appendChild(scrollerEl)
      renderArea.appendChild(editorSlotEl)

      // scroller / 鼠标 / 键盘事件: addEventListener 直接绑, controller 处理
      const sc = scrollerEl
      sc.addEventListener('scroll', (e: Event) => {
        const t = e.target as HTMLElement
        controllerRef.value?.setScroll(t.scrollLeft, t.scrollTop)
      })
      sc.addEventListener('mousedown', (e: MouseEvent) => controllerRef.value?.onMouseDown(e))
      sc.addEventListener('mousemove', (e: MouseEvent) => controllerRef.value?.onMouseMove(e))
      sc.addEventListener('mouseup', (e: MouseEvent) => controllerRef.value?.onMouseUp(e))
      sc.addEventListener('mouseleave', () => controllerRef.value?.onMouseLeave())
      sc.addEventListener('dblclick', (e: MouseEvent) => controllerRef.value?.onDblClick(e))
      sc.addEventListener('contextmenu', (e: MouseEvent) => controllerRef.value?.onContextMenu(e))
      sc.addEventListener('keydown', (e: KeyboardEvent) => controllerRef.value?.onKeyDown(e))

      const controller = new ViewerController(
        {
          canvas: canvasEl,
          renderArea,
          scroller: scrollerEl,
          spacer: spacerEl,
          overlays: { main: ovMainEl, frow: ovFRowEl, fcol: ovFColEl, corner: ovCornerEl },
          editorSlot: editorSlotEl,
        },
        {
          onRenderer: () => {},
          onRenderTick: () => { renderTick.value++ },
          onSelectionChange: () => {
            selVersion.value++
            const sel = controllerRef.value?.getSelection()
            const active = controllerRef.value?.getActiveCell()
            if (sel && active) fire('selection-change', { range: sel, active })
          },
          onCellClick: (row, col, text) => fire('cell-click', { row, col, text }),
          onCellDblClick: (row, col, text) => fire('cell-dblclick', { row, col, text }),
          onHyperlink: (url, cell) => { fire('hyperlink-click', { url, cell }); if (props.openLinks) window.open(url, '_blank', 'noopener') },
          onTooltip: (tip) => { tooltip.value = tip },
          onFindChange: () => { findVersion.value++ },
          onFilterChange: () => { filterVersion.value++ },
          onEditEvent: (event, payload) => fire(event as PluginEvent, payload),
          onContextMenuBefore: (payload) => {
            for (const p of normalizedPlugins.value) {
              if (p.contextMenu) {
                const next = p.contextMenu(payload.ctx, payload.items)
                if (Array.isArray(next)) payload.items.splice(0, payload.items.length, ...next)
              }
            }
            if (props.contextMenu === false) payload.preventDefault()
            fire('before-context-menu' as never, payload)
          },
          onContextMenuShow: (payload) => fire('context-menu' as never, payload),
        },
      )
      controller.fileName = props.fileName
      controller.setEditConfig(effectiveEditConfig.value)
      controller.setEditorResolver(hasEditor.value ? resolveEditor : undefined)
      controller.setLightboxEnabled(props.imageLightbox !== false)
      controller.setContextMenuTransform(typeof props.contextMenu === 'function' ? props.contextMenu : null)
      controllerRef.value = controller
      pluginOverlayHost = new PluginOverlayHost(pluginOvEl)
      initPlugins()
      void runInitialLoad()
      resizeObserver = new ResizeObserver(() => { controllerRef.value?.measure(); controllerRef.value?.render() })
      resizeObserver.observe(renderArea)
      if (typeof document !== 'undefined') document.addEventListener('click', onDocClick)
      })
    })

    // ---- 各 prop 变化 → 同步到 controller ----
    watch(() => [props.src, props.workbook, props.templateFile, runtimeTemplateSrc.value], () => {
      if (controllerRef.value) void runInitialLoad()
    }, { deep: false })
    watch(() => props.fileName, (v) => { const c = controllerRef.value; if (c) c.fileName = v })
    watch(effectiveEditConfig, (cfg) => controllerRef.value?.setEditConfig(cfg))
    watch(() => props.contextMenu, (cm) => controllerRef.value?.setContextMenuTransform(typeof cm === 'function' ? cm : null))
    watch(() => [props.editor, props.plugins], () => controllerRef.value?.setEditorResolver(hasEditor.value ? resolveEditor : undefined), { deep: true })
    watch(() => props.cellImageFit, (fit) => { if (fit) controllerRef.value?.setCellImageFit(fit) })
    watch(() => props.imageLightbox, (v) => controllerRef.value?.setLightboxEnabled(v !== false))
    watch(() => [effectiveTheme.value, props.cellStyle, props.plugins, props.readOnlyCellStyle], () => { if (controllerRef.value) rebuildRenderer() }, { deep: true })
    watch(() => props.plugins, () => initPlugins(), { deep: false })

    watch(workbook, async (wb) => {
      if (!wb) return
      controllerRef.value?.clearFilterState()
      activeSheet.value = wb.activeSheet
      // nextTick 只保证 Vue patch flush, 但 chrome 刚加上, 浏览器 layout/reflow 可能没完成,
      // 此时 renderArea.clientHeight 是中间态. rAF 等到下一帧才 measure → 拿到 final layout 尺寸.
      await nextTick()
      await new Promise<void>((r) => requestAnimationFrame(() => r()))
      rebuildRenderer()
      fire('rendered' as never, wb)
    })
    watch(activeSheet, async (idx, oldIdx) => {
      if (oldIdx != null) controllerRef.value?.resetFilter(workbook.value?.sheets[oldIdx])
      await nextTick()
      rebuildRenderer()
      const wb = workbook.value
      if (wb?.sheets[idx]) fire('sheet-change', { index: idx, name: wb.sheets[idx].name })
    })
    watch(zoom, (z) => controllerRef.value?.setZoom(z))
    watch(error, (msg) => { if (msg) fire('error' as never, msg) })
    watch(progress, (p) => { if (p) fire('progress' as never, p) })
    watch(renderTick, () => renderPluginOverlays())

    onBeforeUnmount(() => {
      resizeObserver?.disconnect()
      controllerRef.value?.dispose()
      pluginOverlayHost?.dispose()
      pluginCleanups.forEach((fn) => fn())
      if (workbook.value) revokeImages(workbook.value)
      if (typeof document !== 'undefined') document.removeEventListener('click', onDocClick)
    })

    // ---- 选区 / 活动格 / 公式栏派生 ----
    const selection = computed<MergeRange | null>(() => { void selVersion.value; return controllerRef.value?.getSelection() ?? null })
    const activeCellAddr = computed(() => {
      void selVersion.value
      const c = controllerRef.value?.getActiveCell()
      return c ? colIndexToLetters(c.col) + (c.row + 1) : ''
    })
    const selRangeLabel = computed(() => {
      const s = selection.value
      if (!s || (s.top === s.bottom && s.left === s.right)) return ''
      return `${colIndexToLetters(s.left)}${s.top + 1}:${colIndexToLetters(s.right)}${s.bottom + 1}`
    })
    const stats = computed(() => {
      void selVersion.value
      const r = controllerRef.value?.renderer
      const s = controllerRef.value?.getSelection() ?? null
      return r && s ? r.selectionStats(s) : null
    })
    function fmtNum(n: number): string {
      if (!isFinite(n)) return '—'
      return n.toLocaleString('en-US', { maximumFractionDigits: 2 })
    }
    const fbDraft = ref('')
    const fbEditing = ref(false)
    // 仅依赖 selVersion (低频). 不读 renderTick — renderTick 每帧 scroll/mousemove ++,
    // 读它会让整个 render function 每帧重跑 → chrome DOM 全部 patch → 滚动卡顿.
    const fbCanEdit = computed(() => { void selVersion.value; return !!controllerRef.value?.canEditActiveCell() })
    const formulaBarEditString = computed(() => { void selVersion.value; return controllerRef.value?.getCellEditString() ?? '' })
    const formulaBarText = computed(() => {
      void selVersion.value
      const r = controllerRef.value?.renderer
      const c = controllerRef.value?.getActiveCell()
      if (!r || !c) return ''
      return r.cellFormula(c.row, c.col) ?? r.cellText(c.row, c.col)
    })
    watch(formulaBarEditString, (v) => { if (!fbEditing.value) fbDraft.value = v }, { immediate: true })
    watch(fbDraft, () => nextTick(syncFbHeight))
    function syncFbHeight() {
      const el = fbSlot.value
      if (!el) return
      el.style.height = 'auto'
      el.style.height = el.scrollHeight + 'px'
    }
    function fbFocus() { fbEditing.value = true; fbDraft.value = formulaBarEditString.value; nextTick(syncFbHeight) }
    function fbCommit(move?: 'down') {
      controllerRef.value?.commitActiveCellValue(fbDraft.value, move)
      fbEditing.value = false
      fbDraft.value = formulaBarEditString.value
      if (move === 'down') scrollerEl?.focus()
    }
    function fbCancel() { fbEditing.value = false; fbDraft.value = formulaBarEditString.value; scrollerEl?.focus() }
    function fbBlur() { if (fbEditing.value) fbCommit() }
    function fbKeydown(e: KeyboardEvent) {
      e.stopPropagation()
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); fbCommit('down') }
      else if (e.key === 'Escape') { e.preventDefault(); fbCancel() }
    }

    // ---- 查找 / 筛选 ----
    const findOpen = ref(false)
    const findState = computed<FindState>(() => { void findVersion.value; return controllerRef.value?.getFindState() ?? { query: '', matchCase: false, wholeCell: false, count: 0, index: -1 } })
    function openFind() { findOpen.value = true }
    function closeFind() { findOpen.value = false; controllerRef.value?.clearFind(); scrollerEl?.focus() }
    function toggleAutoFilter() { controllerRef.value?.toggleAutoFilter() }
    function onRootKeydown(e: KeyboardEvent) {
      if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) { e.preventDefault(); openFind() }
    }

    // ---- 导出 (含进度遮罩链路, 跟 Vue 3 SFC 同语义) ----
    function cancelExport() { exportCtrl?.abort() }
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
    const exportImage = (opts?: ImageExportOptions) => chain(opts, (o) => controllerRef.value!.exportImage(o))
    const downloadImage = (opts?: ImageExportOptions) => chain(opts, (o) => controllerRef.value!.downloadImage(o))
    const exportPdf = (opts?: PdfExportOptions) => chain(opts, (o) => controllerRef.value!.exportPdf(o))
    const downloadPdf = (opts?: PdfExportOptions) => chain(opts, (o) => controllerRef.value!.downloadPdf(o))
    const print = (opts?: PrintOptions) => chain(opts, (o) => controllerRef.value!.print(o))
    async function onExportPdf() { try { await downloadPdf() } catch (e) { reportError(e) } }
    /** 把 ExportDialog 配置映射成各导出方法的入参并执行 (跟 Vue 3 SFC 同语义) */
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
      const page = {
        ...(cfg.format !== 'auto' ? { format: cfg.format } : {}),
        ...(cfg.orientation !== 'auto' ? { orientation: cfg.orientation } : {}),
        fitToWidth: cfg.fitToWidth,
      }
      try {
        if (cfg.action === 'png') await downloadImage(common)
        else if (cfg.action === 'pdf') await downloadPdf({ ...common, ...page, vector: cfg.pdfVector })
        else await print({ ...common, ...page })
      } catch (e) { reportError(e) }
    }
    function reportError(e: unknown) {
      const msg = (e as Error)?.message || String(e)
      console.error('[ooxml-excel-editor/vue2] 导出失败:', e)
      fire('error' as never, msg)
      if (typeof window !== 'undefined') window.alert?.(msg)
    }

    // ---- 模板拾取 ----
    function openTemplateFilePicker() { templateInputSlot.value?.click() }
    function onTemplateFilePicked(e: Event) {
      const f = (e.target as HTMLInputElement).files?.[0]; if (!f) return
      runtimeTemplateSrc.value = f
      runtimeTemplateName.value = f.name
      ;(e.target as HTMLInputElement).value = ''
    }
    function clearRuntimeTemplate() { runtimeTemplateSrc.value = null; runtimeTemplateName.value = null }

    // ---- 数据读取 helper ----
    function dataSheet(sheetIndex?: number): SheetModel | null { const wb = workbook.value; if (!wb) return null; return wb.sheets[sheetIndex ?? activeSheet.value] ?? null }
    function withDate1904<T extends ReadOptions>(opts?: T): T { return { ...(opts as T), date1904: workbook.value?.date1904 ?? false } }

    // ---- 命令式 API ----
    const viewerApi = {
      load: (src: ExcelSource) => load(src, effectiveTransform),
      getWorkbook: () => workbook.value,
      getActiveSheet: () => activeSheet.value,
      setActiveSheet: (i: number) => { if (workbook.value?.sheets[i]) activeSheet.value = i },
      getSelection: () => selection.value,
      setSelection: (range: MergeRange) => controllerRef.value?.setSelectionRange(range),
      rectOf: (row: number, col: number) => controllerRef.value?.rectOf(row, col) ?? null,
      rectOfRange: (range: MergeRange) => controllerRef.value?.rectOfRange(range) ?? null,
      redraw: () => controllerRef.value?.render(),
      isCellEditable: (row: number, col: number) => controllerRef.value?.isCellEditable(row, col) ?? false,
      setEditableTargets: (targets: EditableTarget | EditableTarget[] | undefined) => controllerRef.value?.setEditableTargets(targets),
      getEditableTargets: () => controllerRef.value?.getEditableTargets(),
      editCell: (row: number, col: number, value: any) => controllerRef.value?.editCell(row, col, value) ?? false,
      editRange: (range: MergeRange, values: any[][]) => controllerRef.value?.editRange(range, values) ?? false,
      clearRange: (range: MergeRange) => controllerRef.value?.clearRange(range) ?? false,
      setStyle: (range: MergeRange, patch: CellStyleOverride) => controllerRef.value?.setStyle(range, patch) ?? false,
      getActiveFillColor: () => controllerRef.value?.getActiveFillColor() ?? '#FFFFFF',
      getActiveFontColor: () => controllerRef.value?.getActiveFontColor() ?? '#000000',
      setSelectionFill: (color: string | null) => controllerRef.value?.setSelectionFill(color) ?? false,
      setSelectionFontColor: (color: string) => controllerRef.value?.setSelectionFontColor(color) ?? false,
      getSelectionWrapState: () => controllerRef.value?.getSelectionWrapState() ?? 'none',
      toggleWrapTextOnSelection: () => controllerRef.value?.toggleWrapTextOnSelection() ?? false,
      mergeCells: (range: MergeRange) => controllerRef.value?.mergeCells(range) ?? false,
      unmergeCells: (range: MergeRange) => controllerRef.value?.unmergeCells(range) ?? false,
      pasteText: (text: string, at?: { row: number; col: number }) => controllerRef.value?.pasteText(text, at) ?? false,
      pasteRichHtml: (html: string, at?: { row: number; col: number }) => controllerRef.value?.pasteRichHtml(html, at) ?? false,
      pasteImageBlob: (blob: Blob, at?: { row: number; col: number }) => controllerRef.value?.pasteImageBlob(blob, at) ?? Promise.resolve(false),
      getImages: () => controllerRef.value?.getImages() ?? [],
      addImage: (a: any) => controllerRef.value?.addImage(a) ?? -1,
      removeImage: (i: number) => controllerRef.value?.removeImage(i) ?? false,
      moveImage: (i: number, dx: number, dy: number) => controllerRef.value?.moveImage(i, dx, dy) ?? false,
      resizeImage: (i: number, w: number, h: number) => controllerRef.value?.resizeImage(i, w, h) ?? false,
      getCellEditString: () => controllerRef.value?.getCellEditString() ?? '',
      canEditActiveCell: () => controllerRef.value?.canEditActiveCell() ?? false,
      commitActiveCellValue: (value: string, move?: 'down') => controllerRef.value?.commitActiveCellValue(value, move) ?? false,
      getCellImages: () => controllerRef.value?.getCellImages() ?? [],
      getCellImageAt: (row: number, col: number) => controllerRef.value?.getCellImageAt(row, col) ?? null,
      openImageLightbox: (src: string, fileName?: string, mime?: string) => controllerRef.value?.openImageLightbox(src, fileName, mime),
      setCellImageFit: (fit: 'fill' | 'contain' | 'cover') => controllerRef.value?.setCellImageFit(fit),
      convertImageToCell: (i: number, row: number, col: number) => controllerRef.value?.convertImageToCell(i, row, col) ?? false,
      convertImageToCellAuto: (i: number) => controllerRef.value?.convertImageToCellAuto(i) ?? false,
      convertAllImagesToCells: (col?: number) => controllerRef.value?.convertAllImagesToCells(col) ?? 0,
      convertImagesInRangeToCell: (range: MergeRange) => Promise.resolve(controllerRef.value?.convertImagesInRangeToCell(range) ?? 0),
      convertCellImagesInRangeToFloat: (range: MergeRange, size?: any) => Promise.resolve(controllerRef.value?.convertCellImagesInRangeToFloat(range, size) ?? 0),
      openContextMenu: (x: number, y: number, items?: any[]) => controllerRef.value?.openContextMenu(x, y, items),
      closeContextMenu: () => controllerRef.value?.closeContextMenu(),
      convertCellImageToFloat: (row: number, col: number, size?: any) => controllerRef.value?.convertCellImageToFloat(row, col, size) ?? false,
      insertRows: (at: number, count?: number) => controllerRef.value?.insertRows(at, count) ?? false,
      deleteRows: (at: number, count?: number) => controllerRef.value?.deleteRows(at, count) ?? false,
      insertCols: (at: number, count?: number) => controllerRef.value?.insertCols(at, count) ?? false,
      deleteCols: (at: number, count?: number) => controllerRef.value?.deleteCols(at, count) ?? false,
      undo: () => controllerRef.value?.undo(),
      redo: () => controllerRef.value?.redo(),
      canUndo: () => controllerRef.value?.canUndo() ?? false,
      canRedo: () => controllerRef.value?.canRedo() ?? false,
      getEditingCell: () => controllerRef.value?.getEditingCell() ?? null,
      getCellSnapshot: (row: number, col: number) => controllerRef.value?.getCellSnapshot(row, col) ?? null,
      inspectCell: (row: number, col: number) => controllerRef.value?.inspectCell(row, col) ?? null,
      beginEdit: (row: number, col: number) => controllerRef.value?.beginEdit(row, col) ?? false,
      cancelEdit: () => controllerRef.value?.cancelEdit(),
      isEditing: () => controllerRef.value?.isEditing() ?? false,
      setColumnWidth: (target: any, width: number) => controllerRef.value?.setColumnWidth(target, width) ?? 0,
      setRowHeight: (target: any, height: number) => controllerRef.value?.setRowHeight(target, height) ?? 0,
      autoFitColumns: (target?: any) => controllerRef.value?.autoFitColumns(target) ?? 0,
      autoFitRows: (target?: any) => controllerRef.value?.autoFitRows(target) ?? 0,
      resetColumnWidth: (target: any) => controllerRef.value?.resetColumnWidth(target) ?? 0,
      resetRowHeight: (target: any) => controllerRef.value?.resetRowHeight(target) ?? 0,
      isRecalcReady: () => controllerRef.value?.isRecalcReady() ?? false,
      getVirtualExtent: () => controllerRef.value?.getVirtualExtent() ?? { rows: 0, cols: 0 },
      isDirty: () => controllerRef.value?.isDirty() ?? false,
      resetToOriginal: () => controllerRef.value?.resetToOriginal() ?? false,
      exportImage, downloadImage, exportPdf, downloadPdf, print,
      exportXlsx: (opts?: any) => controllerRef.value!.exportXlsx(opts),
      downloadXlsx: (opts?: any) => controllerRef.value!.downloadXlsx(opts),
      exportJson: (opts?: any) => controllerRef.value?.exportJson(opts) ?? '{}',
      downloadJson: (opts?: any) => controllerRef.value?.downloadJson(opts),
      exportCsv: (opts?: any) => controllerRef.value?.exportCsv(opts) ?? '',
      downloadCsv: (opts?: any) => controllerRef.value?.downloadCsv(opts),
      getCellValue: (row: number, col: number, si?: number) => { const s = dataSheet(si); return s ? getCellValue(s, row, col) : null },
      getCellText: (row: number, col: number, si?: number) => { const s = dataSheet(si); return s ? getCellText(s, row, col, workbook.value?.date1904 ?? false) : '' },
      getSheetData: (opts?: ReadOptions, si?: number) => { const s = dataSheet(si); return s ? getSheetData(s, withDate1904(opts)) : [] },
      getSheetJSON: (opts?: any, si?: number) => { const s = dataSheet(si); return s ? sheetToJSON(s, withDate1904(opts)) : [] },
      getRangeData: (range: MergeRange, opts?: ReadOptions, si?: number) => { const s = dataSheet(si); return s ? getRangeData(s, range, withDate1904(opts)) : [] },
      openTemplateFilePicker,
      clearRuntimeTemplate,
    }
    expose?.(viewerApi)

    // ---- 渲染 helper ----
    function visibleSheets() { const wb = workbook.value; return wb ? wb.sheets.map((s, i) => ({ s, i })).filter(({ s }) => s.state === 'visible') : [] }
    function toggleFreeze() {
      const s = workbook.value?.sheets[activeSheet.value]
      const controller = controllerRef.value
      if (!s || !controller) return
      const fz = s.freeze
      if (fz.frozenRows || fz.frozenCols) s.freeze = { frozenRows: 0, frozenCols: 0 }
      else {
        const c = controller.getActiveCell()
        s.freeze = { frozenRows: c ? c.row : 1, frozenCols: c ? c.col : 0 }
      }
      controller.renderer?.rebuildMetrics()
      controller.refreshContentSize()
      controller.render()
    }

    /** 内置 toolbar item 工厂 — 跟 Vue 3 SFC builtinTool 完全对齐 */
    function bi(o: Partial<ResolvedToolbarItem> & { id: string }): ResolvedToolbarItem {
      return { kind: 'builtin', ...o }
    }
    const I = (name: string) => TOOLBAR_ICONS[name]
    function builtinTool(id: string): ResolvedToolbarItem | null {
      const controller = controllerRef.value
      const sheet = workbook.value?.sheets[activeSheet.value]
      switch (id) {
        case 'find':
          return bi({ id, iconSvg: I('find'), label: '查找', title: '查找 (Ctrl+F)', active: findOpen.value, onClick: () => (findOpen.value ? closeFind() : openFind()) })
        case 'filter':
          return bi({ id, iconSvg: I('filter'), label: '筛选', title: '切换自动筛选', active: !!sheet?.autoFilterRange, onClick: toggleAutoFilter })
        case 'clear-filter':
          return bi({ id, iconSvg: I('clear-filter'), label: '清除筛选', title: '清除当前表全部筛选', disabled: !controller?.hasFilters(), onClick: () => controller?.clearAllFilters() })
        case 'copy':
          return bi({ id, iconSvg: I('copy'), label: '复制', title: '复制选区 (Ctrl+C)', disabled: !selection.value, onClick: () => void controller?.copySelection() })
        case 'wrap-text': {
          const wrapState = controller?.getSelectionWrapState() ?? 'none'
          return bi({ id, iconSvg: I('wrap-text'), label: '自动换行', title: '自动换行(选区,WPS 风格 toggle)', active: wrapState === 'all', disabled: !selection.value || !props.editable, onClick: () => void controller?.toggleWrapTextOnSelection() })
        }
        case 'template': {
          const active = !!effectiveTemplateSrc.value
          const name = effectiveTemplateName.value
          const isXlsxSrc = !!props.src && !props.workbook
          return bi({
            id, iconSvg: I('template'), label: '模板',
            title: isXlsxSrc ? '模板仅对 JSON / 模型数据源生效;当前是 .xlsx 数据源,模板不可用'
              : active ? `模板已加载:${name || '(未命名)'}`
              : '为 JSON / 模型数据源套用 .xlsx 模板的样式;模板的文字内容会被丢弃',
            active, disabled: isXlsxSrc,
            items: [
              bi({ id: 'tpl-default', label: (!active ? '✓ ' : '') + '默认渲染', title: '不套模板,数据按默认样式渲染', disabled: !active, onClick: clearRuntimeTemplate }),
              bi({ id: 'tpl-sep', type: 'separator' }),
              bi({ id: 'tpl-import', label: '导入 .xlsx 模板…', title: '选一份 .xlsx, 把它的 styling 套到当前 JSON 数据上', onClick: openTemplateFilePicker }),
              bi({ id: 'tpl-clear', label: '清除模板', title: '切回默认样式渲染', disabled: !active, onClick: clearRuntimeTemplate }),
            ],
          })
        }
        case 'image-tools': {
          const sel = selection.value
          const active = controller?.getActiveCell()
          const hasFloats = (sheet?.images.length ?? 0) > 0
          return bi({
            id, iconSvg: I('image-tools'), label: '图片工具', title: '浮动图 ⇄ 单元格内嵌图(WPS DISPIMG)互转',
            disabled: !props.editable,
            items: [
              bi({ id: 'img-sel-to-cell', label: '选区:浮动 → 嵌入', title: '把选区里"中心格在选区内"的浮动图就近嵌入', disabled: !sel || !hasFloats, onClick: () => sel && controller?.convertImagesInRangeToCell(sel) }),
              bi({ id: 'img-sel-to-float', label: '选区:嵌入 → 浮动', title: '把选区内所有 DISPIMG 格拎成浮动图', disabled: !sel, onClick: () => sel && controller?.convertCellImagesInRangeToFloat(sel) }),
              bi({ id: 'img-sep', type: 'separator' }),
              bi({ id: 'img-all-to-cell', label: '整表:浮动 → 嵌入', title: '全表浮动图按几何就近嵌入各自单元格', disabled: !hasFloats, onClick: () => controller?.convertAllImagesToCells() }),
              bi({ id: 'img-col-to-cell', label: '整列:浮动 → 嵌入(活动列)', title: '把中心落在活动列的浮动图就近嵌入', disabled: !hasFloats || !active, onClick: () => active && controller?.convertAllImagesToCells(active.col) }),
            ],
          })
        }
        case 'freeze': {
          const fz = sheet?.freeze
          return bi({ id, iconSvg: I('freeze'), label: '冻结', title: '冻结/取消冻结(在活动单元格)', active: !!(fz && (fz.frozenRows || fz.frozenCols)), onClick: toggleFreeze })
        }
        case 'export':
          return bi({
            id, iconSvg: I('export'), label: '导出', title: '导出 / 打印',
            items: [
              bi({ id: 'export-png', label: '导出为图片 (PNG)', onClick: () => void downloadImage() }),
              bi({ id: 'export-pdf', label: '导出为 PDF (位图)', onClick: onExportPdf }),
              bi({ id: 'export-pdf-vector', label: '导出为 PDF (矢量·文字可选)', onClick: () => void downloadPdf({ vector: true }).catch(reportError) }),
              bi({ id: 'export-print', label: '打印…', onClick: () => void print() }),
              bi({ id: 'export-sep', type: 'separator' }),
              bi({ id: 'export-settings', label: '导出设置…', onClick: () => { exportDialogOpen.value = true; exportForm.scope = selRangeLabel.value ? 'selection' : 'sheet' } }),
            ],
          })
        case 'zoom':
          return bi({
            id, iconSvg: I('zoom'), label: Math.round(zoom.value * 100) + '%', title: '缩放',
            items: [50, 75, 100, 125, 150, 200].map((p) => bi({ id: 'zoom-' + p, label: p + '%', active: Math.round(zoom.value * 100) === p, onClick: () => { zoom.value = p / 100 } })),
          })
      }
      return null
    }
    function resolveItem(it: any, kind: 'custom' | 'plugin'): ResolvedToolbarItem {
      const out: ResolvedToolbarItem = { kind, id: it.id, type: it.type, icon: it.icon, label: it.label, title: it.title }
      if (it.active) out.active = !!it.active(viewerApi as ViewerApi)
      if (it.disabled) out.disabled = !!it.disabled(viewerApi as ViewerApi)
      if (it.onClick) out.onClick = () => it.onClick?.(viewerApi as ViewerApi)
      if (it.items?.length) out.items = it.items.map((c: any) => resolveItem(c, kind))
      return out
    }
    const resolvedToolbar = computed<ResolvedToolbarItem[]>(() => {
      void selVersion.value; void findVersion.value; void filterVersion.value
      if (props.toolbar === false) return []
      const entries: Array<string | any> = Array.isArray(props.toolbar) ? props.toolbar : ['find', 'filter']
      const out: ResolvedToolbarItem[] = []
      for (const e of entries) {
        if (typeof e === 'string') {
          if (e === 'separator' || e === '|') out.push({ id: 'sep-' + out.length, type: 'separator', kind: 'builtin' })
          else { const b = builtinTool(e); if (b) out.push(b) }
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

    /** 公共下拉菜单渲染 (用于 header 导出 + action toolbar 各下拉) */
    function renderToolbarMenu(items: ResolvedToolbarItem[], onPick: (it: ResolvedToolbarItem) => void, menuClass = 'ov-tb-menu'): VNode {
      return h('div', { class: menuClass, attrs: { 'data-tb-menu': 'true' } },
        items.map((it) => it.type === 'separator'
          ? h('div', { key: it.id, class: 'sep' })
          : h('button', {
              key: it.id,
              class: { mi: true, active: !!it.active },
              attrs: { disabled: !!it.disabled, title: it.title },
              on: { click: (e: Event) => { e.stopPropagation(); if (!it.disabled) onPick(it) } },
            }, [
              it.iconSvg ? h('span', { class: 'ic', domProps: { innerHTML: svgWrap(it.iconSvg) } })
                : it.icon ? h('span', { class: 'ic-e' }, it.icon)
                : null,
              h('span', { class: 'lb' }, it.label || it.id),
            ]),
        ),
      )
    }

    /** 顶部 ViewerToolbar (1:1 跟 Vue 3 SFC) — 文件名 + 表数 + 导出下拉 + 缩放 [−/+] */
    function renderHeader(): VNode | null {
      if (!workbook.value) return null
      const fname = displayFileName.value
      const tname = effectiveTemplateName.value
      const STEPS = [0.5, 0.75, 1, 1.25, 1.5, 2]
      const exportMenuOpen = menuOpenId.value === '__header-export'
      const exportMenuItems: ResolvedToolbarItem[] = [
        bi({ id: 'h-png', label: '导出为图片 (PNG)', onClick: () => void downloadImage() }),
        bi({ id: 'h-pdf', label: '导出为 PDF (位图)', onClick: onExportPdf }),
        bi({ id: 'h-pdf-vec', label: '导出为 PDF (矢量·文字可选)', onClick: () => void downloadPdf({ vector: true }).catch(reportError) }),
        bi({ id: 'h-print', label: '打印…', onClick: () => void print() }),
        bi({ id: 'h-sep', type: 'separator' }),
        bi({ id: 'h-settings', label: '导出设置…', onClick: () => { exportDialogOpen.value = true; exportForm.scope = selRangeLabel.value ? 'selection' : 'sheet' } }),
      ]
      const setZoom = (z: number) => { zoom.value = Math.min(3, Math.max(0.3, z)) }
      return h('div', { key: 'header', class: 'ov-toolbar' }, [
        h('span', { class: 'file', attrs: { title: fname + (tname ? ' · 模板: ' + tname : '') } }, [
          fname || '未命名工作簿',
          tname ? h('span', { class: 'tpl' }, ' · 模板: ' + tname) : null,
        ]),
        h('span', { class: 'meta' }, visibleSheets().length + ' 个工作表'),
        h('div', { class: 'spacer' }),
        // 导出下拉
        h('div', { class: 'ov-export-wrap', attrs: { 'data-tb-menu': 'true' } }, [
          h('button', {
            class: { 'ov-export-btn': true, open: exportMenuOpen },
            attrs: { title: '导出 / 打印' },
            on: { click: (e: Event) => { e.stopPropagation(); toggleMenu('__header-export') } },
          }, ['导出 ', h('span', { class: 'caret' }, '▾')]),
          exportMenuOpen ? renderToolbarMenu(exportMenuItems, (it) => { setMenuOpen(null); it.onClick?.() }, 'ov-tb-menu header-menu') : null,
        ]),
        // 缩放组 [-/select/+]
        h('div', { class: 'ov-zoom' }, [
          h('button', { attrs: { title: '缩小' }, on: { click: () => setZoom(zoom.value - 0.1) } }, '−'),
          h('select', {
            domProps: { value: String(zoom.value) },
            on: { change: (e: Event) => setZoom(parseFloat((e.target as HTMLSelectElement).value)) },
          }, [
            ...STEPS.map((s) => h('option', { key: s, domProps: { value: String(s) } }, Math.round(s * 100) + '%')),
            !STEPS.includes(zoom.value)
              ? h('option', { key: 'cur', domProps: { value: String(zoom.value) } }, Math.round(zoom.value * 100) + '%')
              : null,
          ]),
          h('button', { attrs: { title: '放大' }, on: { click: () => setZoom(zoom.value + 0.1) } }, '+'),
        ]),
      ])
    }

    /** Action 工具栏 (1:1 跟 Vue 3 SFC ActionToolbar) — SVG 图标 + 下拉子菜单 */
    function renderActionToolbar(): VNode | null {
      if (!workbook.value || !showActionBar.value) return null
      const items = resolvedToolbar.value
      const renderItem = (it: ResolvedToolbarItem): VNode => {
        if (it.type === 'separator') return h('span', { key: it.id, class: 'ov-at-divider' })
        const hasSub = !!it.items?.length
        const open = menuOpenId.value === it.id
        const onBtnClick = (e: Event) => {
          e.stopPropagation()
          if (it.disabled) return
          if (hasSub) toggleMenu(it.id)
          else { it.onClick?.(); setMenuOpen(null) }
        }
        return h('div', { key: it.id, class: 'ov-at-dd', attrs: { 'data-tb-menu': 'true' } }, [
          h('button', {
            class: { 'ov-at-tool': true, active: !!it.active, open },
            attrs: { disabled: !!it.disabled, title: it.title || it.label || it.id },
            on: { click: onBtnClick },
          }, [
            it.iconSvg ? h('span', { class: 'ov-at-ic', domProps: { innerHTML: svgWrap(it.iconSvg) } })
              : it.icon ? h('span', { class: 'ov-at-ic-e' }, it.icon)
              : null,
            it.label ? h('span', { class: 'ov-at-lb' }, it.label) : null,
            hasSub ? h('span', { class: 'ov-at-caret', domProps: { innerHTML: svgWrap(I('caret')) } }) : null,
          ]),
          (hasSub && open) ? renderToolbarMenu(it.items!, (sub) => { setMenuOpen(null); sub.onClick?.() }) : null,
        ])
      }
      return h('div', { key: 'action', class: 'ov-action-toolbar' }, items.map(renderItem))
    }

    function renderFormulaBar(): VNode | null {
      if (!workbook.value) return null
      const editable = fbCanEdit.value
      return h('div', { key: 'fbar', class: 'ov-formula-bar' }, [
        h('span', { class: 'addr' }, activeCellAddr.value || '—'),
        h('span', { class: 'fx' }, 'fx'),
        editable
          ? h('textarea', {
              ref: fbSlot.refName,
              class: 'content content-input',
              attrs: { rows: 1, spellcheck: 'false', title: fbDraft.value },
              domProps: { value: fbDraft.value },
              on: {
                focus: fbFocus, blur: fbBlur, keydown: fbKeydown,
                input: (e: Event) => { fbDraft.value = (e.target as HTMLTextAreaElement).value; syncFbHeight() },
              },
            })
          : h('span', { class: 'content', attrs: { title: formulaBarText.value } }, formulaBarText.value),
      ])
    }

    function renderSheetTabs(): VNode | null {
      const sheets = visibleSheets()
      if (!workbook.value || sheets.length < 2) return null
      return h('div', { key: 'tabs', class: 'ov-sheet-tabs' }, sheets.map(({ s, i }) => h('button', {
        key: i,
        class: { tab: true, active: i === activeSheet.value },
        on: { click: () => { activeSheet.value = i } },
      }, s.name)))
    }

    function renderFindBar(): VNode | null {
      if (!workbook.value || !findOpen.value) return null
      const st = findState.value
      const controller = controllerRef.value
      return h('div', { key: 'findbar', class: 'ov-findbar' }, [
        h('input', {
          attrs: { placeholder: '查找…', autofocus: true },
          domProps: { value: st.query },
          on: {
            input: (e: Event) => controller?.setFindQuery((e.target as HTMLInputElement).value),
            keydown: (e: KeyboardEvent) => {
              if (e.key === 'Enter') e.shiftKey ? controller?.findPrev() : controller?.findNext()
              else if (e.key === 'Escape') closeFind()
            },
          },
        }),
        h('span', { class: 'count' }, st.count ? `${st.index + 1}/${st.count}` : '无结果'),
        h('button', { on: { click: () => controller?.findPrev() } }, '↑'),
        h('button', { on: { click: () => controller?.findNext() } }, '↓'),
        h('button', { on: { click: closeFind } }, '✕'),
      ])
    }

    function renderStatusBar(): VNode | null {
      if (!workbook.value) return null
      const range = selRangeLabel.value || activeCellAddr.value
      const s = stats.value
      const items: VNode[] = []
      items.push(h('span', { class: 'sel' }, range))
      items.push(h('div', { class: 'grow' }))
      if (s && s.numCount > 0) {
        items.push(h('span', '计数 ' + s.count))
        items.push(h('span', '求和 ' + fmtNum(s.sum)))
        items.push(h('span', '平均 ' + fmtNum(s.avg)))
        items.push(h('span', '最大 ' + fmtNum(s.max)))
        items.push(h('span', '最小 ' + fmtNum(s.min)))
      } else if (s && s.count > 0) {
        items.push(h('span', '计数 ' + s.count))
      }
      return h('div', { key: 'statusbar', class: 'ov-status-bar' }, items)
    }

    function renderTooltip(): VNode | null {
      const t = tooltip.value
      if (!t) return null
      return h('div', {
        key: 'tooltip',
        class: 'ov-cell-tooltip ' + (t.kind || ''),
        style: { left: t.x + 'px', top: t.y + 'px' },
      }, t.text)
    }

    /** 内置导出进度遮罩 (P1.5): exportProgress=false 关闭. 居中模态 + stage 标签 + 进度条 + 取消. */
    function renderExportProgress(): VNode | null {
      if (props.exportProgress === false || !exportBusy.value) return null
      const st = exportState.value
      const stageLabel: Record<string, string> = {
        render: '渲染中', compose: '合成中', paginate: '分页中',
        write: '写出文件', zip: 'zip 压缩', convert: '批量转换',
      }
      const ratio = st?.ratio
      const pct = ratio != null ? Math.round(ratio * 100) : null
      const title = st?.label || stageLabel[st?.stage ?? ''] || '处理中…'
      return h('div', { key: 'exp-overlay', class: 'ov-export-progress', attrs: { role: 'dialog', 'aria-modal': 'true' } }, [
        h('div', { class: 'card' }, [
          h('div', { class: 'title' }, title),
          h('div', { class: pct != null ? 'bar' : 'bar indeterminate' }, [
            pct != null ? h('div', { class: 'fill', style: { width: pct + '%' } }) : null,
          ]),
          h('div', { class: 'row' }, [
            h('span', { class: 'pct' }, pct != null ? pct + '%' : '正在处理…'),
            h('button', { class: 'cancel', attrs: { title: '按 Esc 也可取消' }, on: { click: cancelExport } }, '取消'),
          ]),
        ]),
      ])
    }

    /** 导出/打印高级配置对话框 (跟 Vue 3 SFC ExportDialog.vue 同结构) */
    function renderExportDialog(): VNode | null {
      if (!exportDialogOpen.value || !workbook.value) return null
      const sheetCount = visibleSheets().length
      const sel = selection.value
      const hasSel = !!sel && !(sel.top === sel.bottom && sel.left === sel.right)
      const selLabel = sel ? `${colIndexToLetters(sel.left)}${sel.top + 1}:${colIndexToLetters(sel.right)}${sel.bottom + 1}` : ''
      const close = () => { exportDialogOpen.value = false }
      const run = (action: ExportConfig['action']) => { onDialogExport({ ...exportForm, action }) }
      const radio = (name: string, model: keyof ExportConfig, value: any, label: string, hint?: string, disabled?: boolean) => h('label', {
        class: { disabled: !!disabled },
      }, [
        h('input', {
          attrs: { type: 'radio', name, value: String(value), disabled: disabled ? 'disabled' : undefined },
          domProps: { checked: (exportForm as any)[model] === value },
          on: { change: () => { (exportForm as any)[model] = value } },
        }),
        ' ' + label,
        hint ? h('span', { class: 'hint' }, ' ' + hint) : null,
      ])
      const check = (model: keyof ExportConfig, label: string) => h('label', [
        h('input', {
          attrs: { type: 'checkbox' },
          domProps: { checked: (exportForm as any)[model] },
          on: { change: (e: Event) => { (exportForm as any)[model] = (e.target as HTMLInputElement).checked } },
        }),
        ' ' + label,
      ])
      const select = (model: keyof ExportConfig, opts: Array<[any, string]>, numeric = false) => h('select', {
        domProps: { value: String((exportForm as any)[model]) },
        on: { change: (e: Event) => { const v = (e.target as HTMLSelectElement).value; (exportForm as any)[model] = numeric ? Number(v) : v } },
      }, opts.map(([v, lbl]) => h('option', { domProps: { value: String(v) } }, lbl)))
      return h('div', {
        key: 'export-dlg',
        class: 'ov-dlg-mask',
        on: { click: (e: MouseEvent) => { if (e.target === e.currentTarget) close() } },
      }, [
        h('div', { class: 'ov-dlg', attrs: { role: 'dialog', 'aria-label': '导出设置' } }, [
          h('div', { class: 'ov-dlg-head' }, [
            h('span', '导出 / 打印设置'),
            h('button', { class: 'x', attrs: { title: '关闭' }, on: { click: close } }, '×'),
          ]),
          h('div', { class: 'ov-dlg-body' }, [
            h('div', { class: 'field' }, [
              h('label', { class: 'lbl' }, '范围'),
              h('div', { class: 'opts' }, [
                radio('scope', 'scope', 'selection', '当前选区', hasSel ? selLabel : '(未选多格)', !hasSel),
                radio('scope', 'scope', 'sheet', '当前工作表'),
                radio('scope', 'scope', 'all', `全部工作表 (${sheetCount})`),
              ]),
            ]),
            h('div', { class: 'field' }, [
              h('label', { class: 'lbl' }, '清晰度'),
              select('scale', [[1, '标准 (1×)'], [2, '高清 (2×)'], [3, '超清 (3×)']], true),
            ]),
            h('div', { class: 'field' }, [
              h('label', { class: 'lbl' }, '内容'),
              h('div', { class: 'opts inline' }, [check('includeHeaders', '含行列号'), check('gridlines', '网格线')]),
            ]),
            h('div', { class: 'field' }, [
              h('label', { class: 'lbl' }, 'PDF 类型'),
              h('div', { class: 'opts' }, [
                radio('pdfVector', 'pdfVector', false, '位图', '(完整还原观感)'),
                radio('pdfVector', 'pdfVector', true, '矢量', '(文字可选可搜·清晰·文件小;中文需注册字体)'),
              ]),
            ]),
            h('div', { class: 'field' }, [
              h('label', { class: 'lbl' }, [h('span', '纸张'), h('span', { class: 'hint' }, ' (PDF/打印)')]),
              h('div', { class: 'opts inline' }, [
                select('format', [['auto', '自动(跟随表)'], ['a4', 'A4'], ['a3', 'A3'], ['letter', 'Letter']]),
                select('orientation', [['auto', '方向: 自动'], ['portrait', '纵向'], ['landscape', '横向']]),
                check('fitToWidth', '适应页宽'),
              ]),
            ]),
          ]),
          h('div', { class: 'ov-dlg-foot' }, [
            h('button', { class: 'ghost', on: { click: close } }, '取消'),
            h('div', { class: 'grow' }),
            h('button', { on: { click: () => run('png') } }, '导出 PNG'),
            h('button', { on: { click: () => run('pdf') } }, '导出 PDF'),
            h('button', { class: 'primary', on: { click: () => run('print') } }, '打印…'),
          ]),
        ]),
      ])
    }

    /** loading / error / empty 三态浮层 (相对 .ooxml-excel-viewer 主体定位) */
    function renderState(): VNode | null {
      if (loading.value) {
        const pct = progressPct.value
        return h('div', { key: 'state-loading', class: 'ov-state' }, [
          h('div', { class: 'ov-loader' }, [
            h('div', { class: 'ov-loader-label' }, [
              progressLabel.value,
              pct != null ? h('span', ' ' + pct + '%') : null,
            ]),
            h('div', { class: 'ov-loader-track' }, [
              h('div', { class: pct != null ? 'ov-loader-fill' : 'ov-loader-fill indeterminate', style: pct != null ? { width: pct + '%' } : {} }),
            ]),
          ]),
        ])
      }
      if (error.value) return h('div', { key: 'state-error', class: 'ov-state ov-state-error' }, '解析失败: ' + error.value)
      if (!workbook.value) return h('div', { key: 'state-empty', class: 'ov-state ov-state-hint' }, '拖入或选择一个 .xlsx 文件')
      return null
    }

    return () => h('div', {
      class: 'ooxml-excel-viewer',
      on: { keydown: onRootKeydown },
    }, [
      renderHeader(),
      renderActionToolbar(),
      renderFormulaBar(),
      // render-area 外壳 — controller 在 onMounted 时手动 createElement + appendChild 进来,
      // Vue patch 不会动 controller 加的 children. 这里 children 是 Vue 管理的 [user-overlay-slot]:
      // user 通过 v-slot:overlay="{ rectOf, rectOfRange, tick }" 注入自定义 overlay 元素,
      // 跟 Vue 3 SFC 的 #overlay 同语义. user-slot 会被 Vue 渲染在 render-area 第一个 child,
      // controller append 的 canvas/overlays 在后面 (z-index 控制层级).
      // 必须给 key, 否则 Vue 2 patch 没 key 时按 tag 匹配, 会把这个 div 复用成其他 chrome div.
      h('div', { key: 'render-area', ref: renderAreaSlot.refName, class: 'ov-render-area' }, [
        slots.overlay
          ? h('div', { key: 'user-overlay', class: 'ov-user-slot' }, slots.overlay({
              rectOf: (r: number, c: number) => controllerRef.value?.rectOf(r, c) ?? null,
              rectOfRange: (rg: MergeRange) => controllerRef.value?.rectOfRange(rg) ?? null,
              tick: renderTick.value,
            }))
          : null,
      ]),
      renderStatusBar(),
      renderSheetTabs(),
      // findbar / state / tooltip / 导出进度 浮层 — 用 absolute 浮在 viewer 主体上方
      renderFindBar(),
      renderState(),
      renderTooltip(),
      renderExportDialog(),
      renderExportProgress(),
      h('input', {
        key: 'tpl-input',
        ref: templateInputSlot.refName,
        attrs: { type: 'file', accept: '.xlsx,.xlsm', hidden: true },
        on: { change: onTemplateFilePicked },
      }),
    ])
  },
})
