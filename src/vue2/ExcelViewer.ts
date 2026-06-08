/**
 * Vue 2 壳 (Vue 2.7+ Composition API + render function) — 跟 Vue 3 / React 壳同构, 共用同一份 core.
 *
 * 为什么 .ts 而非 .vue?
 *   - 项目同时存在 vue@3 (Vue 3 壳的 peer) + vue@2 (alias `vue2`) → @vitejs/plugin-vue2 解析
 *     @vue/compiler-sfc 时拿到 vue@3 的版本, SFC 编译失败. 改用 render function 完全绕开
 *     SFC 编译路径. 工程上跟 React 壳 (.tsx + hook) 几乎同构.
 *
 * 本文件 ~600 行, 全 features 移植 Vue 3 壳 (props / emits / viewerApi / 头部 / 工具栏 /
 * 公式栏 / sheet 标签 / 查找). 跟 Vue 3 壳的代码组织一一对应, 维护时按章节定位即可.
 */
import Vue, {
  defineComponent,
  h,
  ref,
  computed,
  watch,
  onMounted,
  onBeforeUnmount,
  getCurrentInstance,
  shallowRef,
  nextTick,
  type PropType,
  type VNode,
} from 'vue2'
import type {
  ContextMenuBeforePayload,
  ContextMenuShowPayload,
  ContextMenuTransform,
  ExcelPlugin,
  ExcelPluginContext,
  OverlayContext,
  PermissionDeniedPayload,
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
import type {
  CellChangePayload,
  DimChangePayload,
  DirtyChangePayload,
  ImageChangePayload,
  StructChangePayload,
} from '@/core/edit/edit-controller'
import type { EditorResolver, CellEditorFactory } from '@/core/edit/editor-context'
import { revokeImages } from '@/core/finalize'
import type { ImageExportOptions, PdfExportOptions, PrintOptions } from '@/core/export'
import { ViewerController, type ContextMenuBeforePayload as _CMB, type FindState } from '@/core/viewer/controller'
import { colIndexToLetters } from '@/core/layout/grid-metrics'
import { getCellValue, getCellText, getSheetData, getRangeData, sheetToJSON, type ReadOptions } from '@/core/model/data-access'
import { useExcelDocumentVue2 } from './use-excel-document'
import './excel-viewer.css'

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
  },
  // Vue 2 emits is informational (跟 Vue 3 不一样, 不影响 .emit 行为)
  setup(props, { emit, expose }) {
    const instance = getCurrentInstance()
    const refs = () => (instance?.proxy?.$refs ?? {}) as Record<string, HTMLElement | undefined>

    // ---- 数据加载 + workbook 状态 ----
    const { loading, error, workbook, load, loadModel, progress, sourceBuffer } = useExcelDocumentVue2()
    const activeSheet = ref(0)
    const zoom = ref(1)
    const renderTick = ref(0)
    const selVersion = ref(0)
    const findVersion = ref(0)
    const filterVersion = ref(0)

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
    let controller: ViewerController | null = null
    let resizeObserver: ResizeObserver | null = null
    let pluginOverlayHost: PluginOverlayHost | null = null
    const pluginHandlers = new Map<PluginEvent, Set<(p: any) => void>>()
    let pluginCleanups: Array<() => void> = []

    function fire(event: PluginEvent, payload: any) {
      ;(emit as any)(event, payload)
      pluginHandlers.get(event)?.forEach((h) => h(payload))
    }

    /** 统一加载入口: xlsx 数据源直接 load, JSON 数据源 + 模板 → 解析模板套样式 + loadModel */
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
      const ctx: ExcelPluginContext = { viewer: viewerApi as ViewerApi, on: register, redraw: () => controller?.render() }
      for (const p of normalizedPlugins.value) {
        if (p.events) for (const [ev, fn] of Object.entries(p.events)) if (fn) register(ev as PluginEvent, fn)
        const cleanup = p.setup?.(ctx)
        if (typeof cleanup === 'function') pluginCleanups.push(cleanup)
      }
    }

    function renderPluginOverlays() {
      if (!pluginOverlayHost || !controller) return
      const ctx: OverlayContext = {
        rectOf: (r, c) => controller!.rectOf(r, c),
        rectOfRange: (range) => controller!.rectOfRange(range),
        tick: renderTick.value,
        workbook: workbook.value,
      }
      pluginOverlayHost.render(normalizedPlugins.value, ctx)
    }

    onMounted(() => {
      const r = refs()
      const canvas = r.canvasEl as HTMLCanvasElement | undefined
      const renderArea = r.renderAreaEl, scroller = r.scrollerEl, spacer = r.spacerEl, editorSlot = r.editorSlotEl
      const ovMain = r.ovMainEl, ovFRow = r.ovFRowEl, ovFCol = r.ovFColEl, ovCorner = r.ovCornerEl
      if (!canvas || !renderArea || !scroller || !spacer || !editorSlot || !ovMain || !ovFRow || !ovFCol || !ovCorner) return
      controller = new ViewerController(
        {
          canvas,
          renderArea,
          scroller,
          spacer,
          overlays: { main: ovMain, frow: ovFRow, fcol: ovFCol, corner: ovCorner },
          editorSlot,
        },
        {
          onRenderer: () => {},
          onRenderTick: () => { renderTick.value++ },
          onSelectionChange: () => { selVersion.value++; const sel = controller?.getSelection(); const active = controller?.getActiveCell(); if (sel && active) fire('selection-change', { range: sel, active }) },
          onCellClick: (row, col, text) => fire('cell-click', { row, col, text }),
          onCellDblClick: (row, col, text) => fire('cell-dblclick', { row, col, text }),
          onHyperlink: (url, cell) => { fire('hyperlink-click', { url, cell }); if (props.openLinks) window.open(url, '_blank', 'noopener') },
          onTooltip: () => {},
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
      if (r.pluginOvEl) pluginOverlayHost = new PluginOverlayHost(r.pluginOvEl)
      initPlugins()
      void runInitialLoad()
      resizeObserver = new ResizeObserver(() => { controller?.measure(); controller?.render() })
      resizeObserver.observe(renderArea)
    })

    // ---- 各 prop 变化 → 同步到 controller ----
    watch(() => [props.src, props.workbook, props.templateFile, runtimeTemplateSrc.value], () => { void runInitialLoad() }, { deep: false })
    watch(() => props.fileName, (v) => { if (controller) controller.fileName = v })
    watch(effectiveEditConfig, (cfg) => controller?.setEditConfig(cfg))
    watch(() => props.contextMenu, (cm) => controller?.setContextMenuTransform(typeof cm === 'function' ? cm : null))
    watch(() => [props.editor, props.plugins], () => controller?.setEditorResolver(hasEditor.value ? resolveEditor : undefined), { deep: true })
    watch(() => props.cellImageFit, (fit) => { if (fit) controller?.setCellImageFit(fit) })
    watch(() => props.imageLightbox, (v) => controller?.setLightboxEnabled(v !== false))
    watch(() => [effectiveTheme.value, props.cellStyle, props.plugins, props.readOnlyCellStyle], () => { if (controller) rebuildRenderer() }, { deep: true })
    watch(() => props.plugins, () => initPlugins(), { deep: false })

    watch(workbook, async (wb) => {
      if (!wb) return
      controller?.clearFilterState()
      activeSheet.value = wb.activeSheet
      await nextTick()
      rebuildRenderer()
      fire('rendered' as never, wb)
    })
    watch(activeSheet, async (idx, oldIdx) => {
      if (oldIdx != null) controller?.resetFilter(workbook.value?.sheets[oldIdx])
      await nextTick()
      rebuildRenderer()
      const wb = workbook.value
      if (wb?.sheets[idx]) fire('sheet-change', { index: idx, name: wb.sheets[idx].name })
    })
    watch(zoom, (z) => controller?.setZoom(z))
    watch(error, (msg) => { if (msg) fire('error' as never, msg) })
    watch(progress, (p) => { if (p) fire('progress' as never, p) })
    watch(renderTick, () => renderPluginOverlays())

    onBeforeUnmount(() => {
      resizeObserver?.disconnect()
      controller?.dispose()
      pluginOverlayHost?.dispose()
      pluginCleanups.forEach((fn) => fn())
      if (workbook.value) revokeImages(workbook.value)
    })

    // ---- 选区 / 活动格 / 公式栏派生 ----
    const selection = computed<MergeRange | null>(() => { void selVersion.value; return controller?.getSelection() ?? null })
    const activeCellAddr = computed(() => { void selVersion.value; const c = controller?.getActiveCell(); return c ? colIndexToLetters(c.col) + (c.row + 1) : '' })
    const fbDraft = ref('')
    const fbEditing = ref(false)
    const fbCanEdit = computed(() => { void selVersion.value; void renderTick.value; return !!controller?.canEditActiveCell() })
    const formulaBarEditString = computed(() => { void selVersion.value; void renderTick.value; return controller?.getCellEditString() ?? '' })
    const formulaBarText = computed(() => { void selVersion.value; const r = controller?.renderer; const c = controller?.getActiveCell(); if (!r || !c) return ''; return r.cellFormula(c.row, c.col) ?? r.cellText(c.row, c.col) })
    watch(formulaBarEditString, (v) => { if (!fbEditing.value) fbDraft.value = v }, { immediate: true })
    watch(fbDraft, () => nextTick(syncFbHeight))
    function syncFbHeight() { const el = refs().fbEl as HTMLTextAreaElement | undefined; if (!el) return; el.style.height = 'auto'; el.style.height = el.scrollHeight + 'px' }
    function fbFocus() { fbEditing.value = true; fbDraft.value = formulaBarEditString.value; nextTick(syncFbHeight) }
    function fbCommit(move?: 'down') { controller?.commitActiveCellValue(fbDraft.value, move); fbEditing.value = false; fbDraft.value = formulaBarEditString.value; if (move === 'down') (refs().scrollerEl as HTMLElement | undefined)?.focus() }
    function fbCancel() { fbEditing.value = false; fbDraft.value = formulaBarEditString.value; (refs().scrollerEl as HTMLElement | undefined)?.focus() }
    function fbBlur() { if (fbEditing.value) fbCommit() }
    function fbKeydown(e: KeyboardEvent) {
      e.stopPropagation()
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); fbCommit('down') }
      else if (e.key === 'Escape') { e.preventDefault(); fbCancel() }
    }

    // ---- 查找 / 筛选 ----
    const findOpen = ref(false)
    const findState = computed<FindState>(() => { void findVersion.value; return controller?.getFindState() ?? { query: '', matchCase: false, wholeCell: false, count: 0, index: -1 } })
    function openFind() { findOpen.value = true }
    function closeFind() { findOpen.value = false; controller?.clearFind(); (refs().scrollerEl as HTMLElement | undefined)?.focus() }
    function toggleAutoFilter() { controller?.toggleAutoFilter() }
    function onRootKeydown(e: KeyboardEvent) {
      if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) { e.preventDefault(); openFind() }
    }

    // ---- 导出 (透传到 controller, 跟 Vue 3 / React 一致) ----
    const exportImage = (opts?: ImageExportOptions) => controller!.exportImage(opts)
    const downloadImage = (opts?: ImageExportOptions) => controller!.downloadImage(opts)
    const exportPdf = (opts?: PdfExportOptions) => controller!.exportPdf(opts)
    const downloadPdf = (opts?: PdfExportOptions) => controller!.downloadPdf(opts)
    const print = (opts?: PrintOptions) => controller!.print(opts)
    async function onExportPdf() { try { await downloadPdf() } catch (e) { reportError(e) } }
    function reportError(e: unknown) { const msg = (e as Error)?.message || String(e); console.error('[ooxml-excel-editor/vue2] 导出失败:', e); fire('error' as never, msg); if (typeof window !== 'undefined') window.alert?.(msg) }

    // ---- 模板拾取器 ----
    function openTemplateFilePicker() { (refs().templateInputEl as HTMLInputElement | undefined)?.click() }
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

    // ---- 命令式 API (跟 Vue 3 viewerApi / React Handle 完全对齐) ----
    const viewerApi = {
      load: (src: ExcelSource) => load(src, effectiveTransform),
      getWorkbook: () => workbook.value,
      getActiveSheet: () => activeSheet.value,
      setActiveSheet: (i: number) => { if (workbook.value?.sheets[i]) activeSheet.value = i },
      getSelection: () => selection.value,
      setSelection: (range: MergeRange) => controller?.setSelectionRange(range),
      rectOf: (row: number, col: number) => controller?.rectOf(row, col) ?? null,
      rectOfRange: (range: MergeRange) => controller?.rectOfRange(range) ?? null,
      redraw: () => controller?.render(),
      isCellEditable: (row: number, col: number) => controller?.isCellEditable(row, col) ?? false,
      setEditableTargets: (targets: EditableTarget | EditableTarget[] | undefined) => controller?.setEditableTargets(targets),
      getEditableTargets: () => controller?.getEditableTargets(),
      editCell: (row: number, col: number, value: any) => controller?.editCell(row, col, value) ?? false,
      editRange: (range: MergeRange, values: any[][]) => controller?.editRange(range, values) ?? false,
      clearRange: (range: MergeRange) => controller?.clearRange(range) ?? false,
      setStyle: (range: MergeRange, patch: CellStyleOverride) => controller?.setStyle(range, patch) ?? false,
      getActiveFillColor: () => controller?.getActiveFillColor() ?? '#FFFFFF',
      getActiveFontColor: () => controller?.getActiveFontColor() ?? '#000000',
      setSelectionFill: (color: string | null) => controller?.setSelectionFill(color) ?? false,
      setSelectionFontColor: (color: string) => controller?.setSelectionFontColor(color) ?? false,
      getSelectionWrapState: () => controller?.getSelectionWrapState() ?? 'none',
      toggleWrapTextOnSelection: () => controller?.toggleWrapTextOnSelection() ?? false,
      mergeCells: (range: MergeRange) => controller?.mergeCells(range) ?? false,
      unmergeCells: (range: MergeRange) => controller?.unmergeCells(range) ?? false,
      pasteText: (text: string, at?: { row: number; col: number }) => controller?.pasteText(text, at) ?? false,
      pasteRichHtml: (html: string, at?: { row: number; col: number }) => controller?.pasteRichHtml(html, at) ?? false,
      pasteImageBlob: (blob: Blob, at?: { row: number; col: number }) => controller?.pasteImageBlob(blob, at) ?? Promise.resolve(false),
      getImages: () => controller?.getImages() ?? [],
      addImage: (a: any) => controller?.addImage(a) ?? -1,
      removeImage: (i: number) => controller?.removeImage(i) ?? false,
      moveImage: (i: number, dx: number, dy: number) => controller?.moveImage(i, dx, dy) ?? false,
      resizeImage: (i: number, w: number, h: number) => controller?.resizeImage(i, w, h) ?? false,
      getCellEditString: () => controller?.getCellEditString() ?? '',
      canEditActiveCell: () => controller?.canEditActiveCell() ?? false,
      commitActiveCellValue: (value: string, move?: 'down') => controller?.commitActiveCellValue(value, move) ?? false,
      getCellImages: () => controller?.getCellImages() ?? [],
      getCellImageAt: (row: number, col: number) => controller?.getCellImageAt(row, col) ?? null,
      openImageLightbox: (src: string, fileName?: string, mime?: string) => controller?.openImageLightbox(src, fileName, mime),
      setCellImageFit: (fit: 'fill' | 'contain' | 'cover') => controller?.setCellImageFit(fit),
      convertImageToCell: (i: number, row: number, col: number) => controller?.convertImageToCell(i, row, col) ?? false,
      convertImageToCellAuto: (i: number) => controller?.convertImageToCellAuto(i) ?? false,
      convertAllImagesToCells: (col?: number) => controller?.convertAllImagesToCells(col) ?? 0,
      convertImagesInRangeToCell: (range: MergeRange) => Promise.resolve(controller?.convertImagesInRangeToCell(range) ?? 0),
      convertCellImagesInRangeToFloat: (range: MergeRange, size?: any) => Promise.resolve(controller?.convertCellImagesInRangeToFloat(range, size) ?? 0),
      openContextMenu: (x: number, y: number, items?: any[]) => controller?.openContextMenu(x, y, items),
      closeContextMenu: () => controller?.closeContextMenu(),
      convertCellImageToFloat: (row: number, col: number, size?: any) => controller?.convertCellImageToFloat(row, col, size) ?? false,
      insertRows: (at: number, count?: number) => controller?.insertRows(at, count) ?? false,
      deleteRows: (at: number, count?: number) => controller?.deleteRows(at, count) ?? false,
      insertCols: (at: number, count?: number) => controller?.insertCols(at, count) ?? false,
      deleteCols: (at: number, count?: number) => controller?.deleteCols(at, count) ?? false,
      undo: () => controller?.undo(),
      redo: () => controller?.redo(),
      canUndo: () => controller?.canUndo() ?? false,
      canRedo: () => controller?.canRedo() ?? false,
      getEditingCell: () => controller?.getEditingCell() ?? null,
      getCellSnapshot: (row: number, col: number) => controller?.getCellSnapshot(row, col) ?? null,
      inspectCell: (row: number, col: number) => controller?.inspectCell(row, col) ?? null,
      beginEdit: (row: number, col: number) => controller?.beginEdit(row, col) ?? false,
      cancelEdit: () => controller?.cancelEdit(),
      isEditing: () => controller?.isEditing() ?? false,
      setColumnWidth: (target: any, width: number) => controller?.setColumnWidth(target, width) ?? 0,
      setRowHeight: (target: any, height: number) => controller?.setRowHeight(target, height) ?? 0,
      autoFitColumns: (target?: any) => controller?.autoFitColumns(target) ?? 0,
      autoFitRows: (target?: any) => controller?.autoFitRows(target) ?? 0,
      resetColumnWidth: (target: any) => controller?.resetColumnWidth(target) ?? 0,
      resetRowHeight: (target: any) => controller?.resetRowHeight(target) ?? 0,
      isRecalcReady: () => controller?.isRecalcReady() ?? false,
      getVirtualExtent: () => controller?.getVirtualExtent() ?? { rows: 0, cols: 0 },
      isDirty: () => controller?.isDirty() ?? false,
      resetToOriginal: () => controller?.resetToOriginal() ?? false,
      exportImage, downloadImage, exportPdf, downloadPdf, print,
      exportXlsx: (opts?: any) => controller!.exportXlsx(opts),
      downloadXlsx: (opts?: any) => controller!.downloadXlsx(opts),
      exportJson: (opts?: any) => controller?.exportJson(opts) ?? '{}',
      downloadJson: (opts?: any) => controller?.downloadJson(opts),
      exportCsv: (opts?: any) => controller?.exportCsv(opts) ?? '',
      downloadCsv: (opts?: any) => controller?.downloadCsv(opts),
      getCellValue: (row: number, col: number, si?: number) => { const s = dataSheet(si); return s ? getCellValue(s, row, col) : null },
      getCellText: (row: number, col: number, si?: number) => { const s = dataSheet(si); return s ? getCellText(s, row, col, workbook.value?.date1904 ?? false) : '' },
      getSheetData: (opts?: ReadOptions, si?: number) => { const s = dataSheet(si); return s ? getSheetData(s, withDate1904(opts)) : [] },
      getSheetJSON: (opts?: any, si?: number) => { const s = dataSheet(si); return s ? sheetToJSON(s, withDate1904(opts)) : [] },
      getRangeData: (range: MergeRange, opts?: ReadOptions, si?: number) => { const s = dataSheet(si); return s ? getRangeData(s, range, withDate1904(opts)) : [] },
    }
    expose?.(viewerApi)

    // ---- 渲染 ----
    function onScroll(e: Event) { const sc = e.target as HTMLElement; controller?.setScroll(sc.scrollLeft, sc.scrollTop) }
    function visibleSheets() { const wb = workbook.value; return wb ? wb.sheets.map((s, i) => ({ s, i })).filter(({ s }) => s.state === 'visible') : [] }
    function toggleFreeze() { const s = workbook.value?.sheets[activeSheet.value]; if (!s || !controller) return; const fz = s.freeze; if (fz.frozenRows || fz.frozenCols) s.freeze = { frozenRows: 0, frozenCols: 0 }; else { const c = controller.getActiveCell(); s.freeze = { frozenRows: c ? c.row : 1, frozenCols: c ? c.col : 0 } }; controller.renderer?.rebuildMetrics(); controller.refreshContentSize(); controller.render() }

    /** 渲染头部 (文件名 + 模板名 + 工具栏导出) */
    function renderHeader(): VNode | null {
      if (!workbook.value) return null
      const fname = displayFileName.value
      const tname = effectiveTemplateName.value
      return h('div', { class: 'ov-toolbar' }, [
        h('span', { class: 'file', attrs: { title: fname + (tname ? ' · 模板: ' + tname : '') } }, [
          h('span', { class: 'name' }, fname || '未命名'),
          tname ? h('span', { class: 'tpl' }, ' · 模板: ' + tname) : null,
          h('span', { class: 'sheets' }, ' · ' + visibleSheets().length + ' 表'),
        ]),
        h('span', { class: 'spacer' }),
        h('select', {
          attrs: { title: '缩放' },
          domProps: { value: Math.round(zoom.value * 100) },
          on: { change: (e: Event) => { zoom.value = Number((e.target as HTMLSelectElement).value) / 100 } },
        }, [50, 75, 100, 125, 150, 200].map((p) => h('option', { domProps: { value: p } }, p + '%'))),
        h('button', { class: 'btn', on: { click: () => void downloadImage() } }, '↓PNG'),
        h('button', { class: 'btn', on: { click: onExportPdf } }, '↓PDF'),
        h('button', { class: 'btn', on: { click: () => void viewerApi.downloadXlsx() } }, '↓XLSX'),
      ])
    }

    /** 渲染 action toolbar (查找/筛选/冻结 + 插件项) */
    function renderActionToolbar(): VNode | null {
      if (!workbook.value || props.toolbar === false) return null
      void renderTick.value; void selVersion.value; void findVersion.value; void filterVersion.value
      const items: VNode[] = []
      items.push(h('button', { class: { tool: true, active: findOpen.value }, attrs: { title: '查找 Ctrl+F' }, on: { click: () => findOpen.value ? closeFind() : openFind() } }, '🔍 查找'))
      items.push(h('button', { class: { tool: true, active: !!workbook.value.sheets[activeSheet.value]?.autoFilterRange }, attrs: { title: '切换自动筛选' }, on: { click: toggleAutoFilter } }, '⏷ 筛选'))
      items.push(h('button', { class: 'tool', attrs: { title: '清除筛选', disabled: !controller?.hasFilters() }, on: { click: () => controller?.clearAllFilters() } }, '✕ 筛选'))
      items.push(h('button', { class: 'tool', attrs: { title: '复制选区', disabled: !selection.value }, on: { click: () => void controller?.copySelection() } }, '⎘'))
      items.push(h('button', { class: { tool: true, active: !!(workbook.value.sheets[activeSheet.value]?.freeze.frozenRows || workbook.value.sheets[activeSheet.value]?.freeze.frozenCols) }, attrs: { title: '冻结活动单元格' }, on: { click: toggleFreeze } }, '❄ 冻结'))
      // 插件贡献工具栏项
      for (const p of normalizedPlugins.value) {
        for (const it of p.toolbar ?? []) {
          if (it.type === 'separator') continue
          items.push(h('button', {
            class: { tool: true, active: !!it.active?.(viewerApi as ViewerApi) },
            attrs: { title: it.title, disabled: !!it.disabled?.(viewerApi as ViewerApi) },
            on: { click: () => it.onClick?.(viewerApi as ViewerApi) },
          }, [it.label ?? it.icon ?? it.id]))
        }
      }
      return h('div', { class: 'ov-action-toolbar' }, items)
    }

    /** 公式栏 textarea + auto-resize (Phase 1.2.1 撑高) */
    function renderFormulaBar(): VNode | null {
      if (!workbook.value) return null
      const editable = fbCanEdit.value
      return h('div', { class: 'ov-formula-bar' }, [
        h('span', { class: 'addr' }, activeCellAddr.value || '—'),
        h('span', { class: 'fx' }, 'fx'),
        editable
          ? h('textarea', {
              ref: 'fbEl',
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

    /** sheet 标签 (多表切换) */
    function renderSheetTabs(): VNode | null {
      const sheets = visibleSheets()
      if (!workbook.value || sheets.length < 2) return null
      return h('div', { class: 'ov-sheet-tabs' }, sheets.map(({ s, i }) => h('button', {
        key: i,
        class: { tab: true, active: i === activeSheet.value },
        on: { click: () => { activeSheet.value = i } },
      }, s.name)))
    }

    /** 查找条 (Ctrl+F) */
    function renderFindBar(): VNode | null {
      if (!workbook.value || !findOpen.value) return null
      const st = findState.value
      return h('div', { class: 'ov-findbar' }, [
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

    return () => h('div', {
      ref: 'rootEl',
      class: 'ooxml-excel-viewer',
      on: { keydown: onRootKeydown },
    }, [
      renderHeader(),
      renderActionToolbar(),
      renderFormulaBar(),
      h('div', { ref: 'renderAreaEl', class: 'ov-render-area' }, [
        h('canvas', { ref: 'canvasEl', class: 'ov-grid-canvas' }),
        h('div', { ref: 'ovMainEl', class: 'ov ov-main' }),
        h('div', { ref: 'ovFColEl', class: 'ov ov-fcol' }),
        h('div', { ref: 'ovFRowEl', class: 'ov ov-frow' }),
        h('div', { ref: 'ovCornerEl', class: 'ov ov-corner' }),
        h('div', { class: 'ov ov-plugin' }, [h('div', { ref: 'pluginOvEl' })]),
        h('div', {
          ref: 'scrollerEl',
          class: 'ov-scroller',
          attrs: { tabindex: '0' },
          on: {
            scroll: onScroll,
            mousedown: (e: MouseEvent) => controller?.onMouseDown(e),
            mousemove: (e: MouseEvent) => controller?.onMouseMove(e),
            mouseup: (e: MouseEvent) => controller?.onMouseUp(e),
            mouseleave: () => controller?.onMouseLeave(),
            dblclick: (e: MouseEvent) => controller?.onDblClick(e),
            contextmenu: (e: MouseEvent) => controller?.onContextMenu(e),
            keydown: (e: KeyboardEvent) => controller?.onKeyDown(e),
          },
        }, [h('div', { ref: 'spacerEl', class: 'ov-spacer' })]),
        h('div', { ref: 'editorSlotEl', class: 'ov-editor-slot' }),
        renderFindBar(),
      ]),
      renderSheetTabs(),
      // 隐藏文件拾取器 (用于工具栏 / 命令式触发 setRuntimeTemplate)
      h('input', {
        ref: 'templateInputEl',
        attrs: { type: 'file', accept: '.xlsx,.xlsm', hidden: true },
        on: { change: onTemplateFilePicked },
      }),
    ])
  },
})
