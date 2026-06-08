/**
 * React 版 ExcelViewer —— 与 Vue 版共用同一套框架无关 core(ViewerController + 引擎)。
 * 本壳只做: DOM 容器 + 把 ViewerController 接到 React 生命周期 + chrome(工具栏/公式栏/状态栏/标签/查找/筛选)。
 * 渲染/选区/交互/查找/筛选/导出 全部由 core 完成,React 与 Vue 共享 ~100% 引擎代码。
 */
import {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useReducer,
  useRef,
  useState,
  type CSSProperties,
} from 'react'
import type { CellModel, CellStyleFn, CellStyleOverride, ImageAnchor, MergeRange, SheetModel, TransformModelFn, WorkbookModel } from '@/core/model/types'
import type { EditableTarget, EditConfig } from '@/core/edit/types'
import type { FormulaEngineFactory } from '@/core/formula/engine'
import type { CellChangePayload, DimChangePayload, DirtyChangePayload, ImageChangePayload, StructChangePayload } from '@/core/edit/edit-controller'
import type { CellSnapshot } from '@/core/model/snapshot'
import type { CellInspection } from '@/core/model/inspect'
import { applyStyleTemplate } from '@/core/template/style-overlay'
import { loadArrayBuffer } from '@/core/loader'
import { detectFormat, finalizeImages } from '@/core/finalize'
import { parseInWorker } from '@/composables/worker-client'
import { jsonToWorkbook, isWorkbookModel, type JsonInput, type JsonLoadOptions } from '@/core/loader-json'
import type { ExportProgress } from '@/core/progress'
import { ExportProgressOverlay } from './ExportProgressOverlay'
import type { CellValue } from '@/core/model/data-access'
import type { EditorResolver, CellEditorFactory } from '@/core/edit/editor-context'
import type { ViewerTheme } from '@/core/render/theme'
import type { ExcelSource } from '@/core/loader'
import type { ImageExportOptions, PdfExportOptions, PrintOptions } from '@/core/export'
import type { XlsxExportOptions } from '@/core/export/xlsx-writer'
import {
  getCellValue,
  getCellText,
  getSheetData,
  getRangeData,
  sheetToJSON,
  type ReadOptions,
  type SheetToJSONOptions,
} from '@/core/model/data-access'
import { colIndexToLetters } from '@/core/layout/grid-metrics'
import { ViewerController, type Cell, type TooltipState } from '@/core/viewer/controller'
import { PluginOverlayHost } from '@/core/viewer/plugin-overlay'
import type {
  ContextMenuBeforePayload,
  ContextMenuShowPayload,
  ContextMenuTransform,
  ExcelPlugin,
  MenuItem,
  OverlayContext,
  PluginEvent,
  ViewerApi,
} from '@/core/plugin'
import { useExcelDocument } from './use-excel-document'
import './excel-viewer.css'

export interface ExcelViewerProps {
  src?: ExcelSource
  /**
   * 直接喂 WorkbookModel 或 JsonInput(P3):绕过 parser。优先级 workbook > src。
   * WorkbookModel 直用;JsonInput(二维数组 / 对象数组 / `{sheets:[...]}`)走 jsonToWorkbook。
   */
  workbook?: WorkbookModel | JsonInput
  /** JSON 直渲选项(workbook = JsonInput 时生效) */
  jsonOptions?: JsonLoadOptions
  /**
   * 渲染模板(P3 重设计 2026-06-08):一份 .xlsx 当**样式捐赠者** —— 模板的 styling
   * (styles / merges / 列宽 / 行高 / freeze / theme)套到无格式数据源上,模板的 raw 文字 / 占位符 /
   * 图 / 图表 / 条件格式 全部丢弃。
   *
   * ⚠️ **只在数据源是 `workbook`(JSON / 模型)时生效**;`src`(.xlsx)数据源自带格式,
   * 给 `templateFile` 会被忽略并 console.warn.
   */
  templateFile?: ExcelSource
  /** 模板显示名(标题栏 `· 模板: xxx` 后缀);不给则取运行时 File.name */
  templateName?: string
  /**
   * 内置导出进度遮罩(P1.5):默认 `true` —— 调 `viewer.downloadPdf/exportImage/...` /
   * 选区图片批量转换 时,壳自动建 AbortController + 接 onProgress → 显示居中模态 + 取消。
   * `false` 关闭(走纯回调);`renderExportProgress` 自渲染(覆盖内置 UI)。
   */
  exportProgress?: boolean
  /** 完全自渲染遮罩:返回任意 ReactNode 替代内置 UI。`exportProgress=false` 时此项也不渲染。 */
  renderExportProgress?: (ctx: { state: ExportProgress | null; busy: boolean; cancel: () => void }) => React.ReactNode
  /**
   * 右键菜单(Plan C):
   * - `false` → 不弹内置菜单(`onBeforeContextMenu` / `onContextMenuShow` 仍触发,壳可自渲染)
   * - 函数 `(ctx, items) => MenuItem[] | undefined` → 在内置 items 上加 / 减 / 重排;返 undefined 不动
   * - 不传(默认)→ editable 时显示内置菜单,非 editable 走浏览器默认
   */
  contextMenu?: boolean | ContextMenuTransform
  /** 右键菜单触发前:`payload.preventDefault()` 阻止内置菜单(自渲染替代) */
  onBeforeContextMenu?: (payload: ContextMenuBeforePayload) => void
  /** 右键菜单展示通知:无论内置是否弹都触发,供自渲染或事件流串到业务 */
  onContextMenuShow?: (payload: ContextMenuShowPayload) => void
  fileName?: string
  theme?: Partial<ViewerTheme>
  /** 单击超链接是否自动打开(默认 true) */
  openLinks?: boolean
  transformModel?: TransformModelFn
  cellStyle?: CellStyleFn
  /** WPS 单元格内嵌图(DISPIMG)贴合方式:contain 等比(默认,与 WPS 渲染一致)/ fill 拉伸铺满 / cover 等比裁剪 */
  cellImageFit?: 'fill' | 'contain' | 'cover'
  /** 图片点击放大灯箱(默认 true):只读模式单击图放大、编辑模式右键「查看大图」;false 关闭 */
  imageLightbox?: boolean
  /** 插件(与 Vue 通用):theme/transformModel/cellStyle/events/overlay/setup 全跨框架可用 */
  plugins?: ExcelPlugin[]
  /** 编辑总开关:默认 false = 只读(行为不变)。开启后才能进入编辑(E0:闸门) */
  editable?: boolean
  /** 按格只读判定:返回 true = 只读(cell 为空格时传 null) */
  cellReadOnly?: (cell: CellModel | null, pos: { row: number; col: number }) => boolean | void
  /** 只读区域(0-based 闭区间);命中即只读 */
  readOnlyRanges?: MergeRange[]
  /**
   * **可编辑白名单**(2026-06-08 新增):设了就是白名单语义,只命中**任一** target 的格才可编辑.
   * 4 种 target 形状自动识别:`{row,col}` 单格 / `{row}` 整行 / `{col}` 整列 / `MergeRange` 矩形.
   * 单值或数组都支持;允许**不相邻**多 target. `undefined`=默认全可编辑(老行为),`[]`=全只读.
   */
  editableTargets?: EditableTarget | EditableTarget[]
  /** 自定义单元格编辑器(按格返回工厂;覆盖插件 editor)。需 editable 开启 */
  editor?: EditorResolver
  /** 公式重算(E4):默认 false 沿用缓存值。开启后编辑公式/被引用格 → 依赖格自动重算。需 editable */
  recalc?: boolean
  /** 自定义/自研公式引擎工厂(可换引擎);不给则用默认 HyperFormula(需 npm i hyperformula) */
  formulaEngine?: FormulaEngineFactory
  className?: string
  style?: CSSProperties
  onRendered?: (wb: WorkbookModel) => void
  onError?: (msg: string) => void
  onCellClick?: (p: { row: number; col: number; text: string }) => void
  onCellDblClick?: (p: { row: number; col: number; text: string }) => void
  onSelectionChange?: (p: { range: MergeRange; active: Cell }) => void
  onHyperlinkClick?: (p: { url: string; cell: Cell }) => void
  onSheetChange?: (p: { index: number; name: string }) => void
  /** 单元格变更(编辑/撤销/重做;含前后完整快照) */
  onCellChange?: (p: CellChangePayload) => void
  onEditStart?: (p: unknown) => void
  onEditCommit?: (p: unknown) => void
  /** 列宽/行高变更(拖拽/autofit/API/撤销重做;前后 px 尺寸) */
  onDimChange?: (p: DimChangePayload) => void
  /** 脏状态变更(有/无未保存修改) */
  onDirtyChange?: (p: DirtyChangePayload) => void
  /** 图片增删移改(前后 ImageAnchor) */
  onImageChange?: (p: ImageChangePayload) => void
  /** 行列结构变更(增删行列) */
  onStructChange?: (p: StructChangePayload) => void
}

/** 命令式句柄(与 Vue ref / ViewerApi 对齐) */
export interface ExcelViewerHandle {
  load: (src: ExcelSource) => void
  getWorkbook: () => WorkbookModel | null
  getActiveSheet: () => number
  setActiveSheet: (i: number) => void
  getSelection: () => MergeRange | null
  setSelection: (range: MergeRange) => void
  rectOf: (row: number, col: number) => { x: number; y: number; w: number; h: number } | null
  rectOfRange: (range: MergeRange) => { x: number; y: number; w: number; h: number } | null
  redraw: () => void
  isCellEditable: (row: number, col: number) => boolean
  setEditableTargets: (targets: EditableTarget | EditableTarget[] | undefined) => void
  getEditableTargets: () => EditableTarget | EditableTarget[] | undefined
  editCell: (row: number, col: number, value: CellValue) => boolean
  editRange: (range: MergeRange, values: CellValue[][]) => boolean
  clearRange: (range: MergeRange) => boolean
  setStyle: (range: MergeRange, patch: CellStyleOverride) => boolean
  getActiveFillColor: () => string
  getActiveFontColor: () => string
  setSelectionFill: (color: string | null) => boolean
  setSelectionFontColor: (color: string) => boolean
  getSelectionWrapState: () => 'all' | 'none' | 'mixed'
  toggleWrapTextOnSelection: () => boolean
  mergeCells: (range: MergeRange) => boolean
  unmergeCells: (range: MergeRange) => boolean
  pasteText: (text: string, at?: { row: number; col: number }) => boolean
  pasteRichHtml: (html: string, at?: { row: number; col: number }) => boolean
  pasteImageBlob: (blob: Blob, at?: { row: number; col: number }) => Promise<boolean>
  getImages: () => ImageAnchor[]
  addImage: (anchor: ImageAnchor) => number
  removeImage: (index: number) => boolean
  moveImage: (index: number, dxPx: number, dyPx: number) => boolean
  resizeImage: (index: number, widthPx: number, heightPx: number) => boolean
  getCellImages: () => { id: string; src: string; mime?: string }[]
  getCellImageAt: (row: number, col: number) => { id: string; src: string; mime?: string } | null
  openImageLightbox: (src: string, fileName?: string, mime?: string) => void
  /** 活动格在公式栏里的可编辑字符串(公式→=…,数值→原始数字串) */
  getCellEditString: () => string
  /** 活动格此刻是否可经公式栏编辑(editable + 非只读) */
  canEditActiveCell: () => boolean
  /** 经公式栏提交活动格的值(move='down' 提交后下移) */
  commitActiveCellValue: (value: string, move?: 'down') => boolean
  setCellImageFit: (fit: 'fill' | 'contain' | 'cover') => void
  convertImageToCell: (imageIndex: number, row: number, col: number) => boolean
  convertImageToCellAuto: (imageIndex: number) => boolean
  convertAllImagesToCells: (col?: number) => number
  convertImagesInRangeToCell: (range: MergeRange) => Promise<number>
  convertCellImagesInRangeToFloat: (range: MergeRange, size?: { width: number; height: number }) => Promise<number>
  openContextMenu: (x: number, y: number, items?: MenuItem[]) => void
  closeContextMenu: () => void
  convertCellImageToFloat: (row: number, col: number, size?: { width: number; height: number }) => boolean
  insertRows: (at: number, count?: number) => boolean
  deleteRows: (at: number, count?: number) => boolean
  insertCols: (at: number, count?: number) => boolean
  deleteCols: (at: number, count?: number) => boolean
  undo: () => void
  redo: () => void
  canUndo: () => boolean
  canRedo: () => boolean
  getEditingCell: () => { row: number; col: number } | null
  getCellSnapshot: (row: number, col: number) => CellSnapshot | null
  inspectCell: (row: number, col: number) => CellInspection | null
  beginEdit: (row: number, col: number) => boolean
  cancelEdit: () => void
  isEditing: () => boolean
  setColumnWidth: (col: number, width: number) => boolean
  setRowHeight: (row: number, height: number) => boolean
  isRecalcReady: () => boolean
  getVirtualExtent: () => { rows: number; cols: number }
  isDirty: () => boolean
  resetToOriginal: () => boolean
  exportImage: (opts?: ImageExportOptions) => Promise<Blob>
  downloadImage: (opts?: ImageExportOptions) => Promise<void>
  exportPdf: (opts?: PdfExportOptions) => Promise<Blob>
  downloadPdf: (opts?: PdfExportOptions) => Promise<void>
  print: (opts?: PrintOptions) => Promise<void>
  exportXlsx: (opts?: XlsxExportOptions) => Promise<Blob>
  downloadXlsx: (opts?: XlsxExportOptions) => Promise<void>
  exportJson: (opts?: SheetToJSONOptions) => string
  downloadJson: (opts?: SheetToJSONOptions) => void
  exportCsv: (opts?: { target?: number; format?: boolean }) => string
  downloadCsv: (opts?: { target?: number; format?: boolean }) => void
  getCellValue: (row: number, col: number, sheet?: number) => ReturnType<typeof getCellValue>
  getCellText: (row: number, col: number, sheet?: number) => string
  getSheetData: (opts?: ReadOptions, sheet?: number) => ReturnType<typeof getSheetData>
  getSheetJSON: (opts?: Parameters<typeof sheetToJSON>[1], sheet?: number) => ReturnType<typeof sheetToJSON>
  getRangeData: (range: MergeRange, opts?: ReadOptions, sheet?: number) => ReturnType<typeof getRangeData>
}

function fmtNum(n: number): string {
  if (!isFinite(n)) return '—'
  return n.toLocaleString('en-US', { maximumFractionDigits: 2 })
}

export const ExcelViewer = forwardRef<ExcelViewerHandle, ExcelViewerProps>(function ExcelViewer(props, ref) {
  const { loading, error, workbook, progress, load, loadModel, sourceBuffer } = useExcelDocument()
  const resolveWb = (w?: WorkbookModel | JsonInput): WorkbookModel | null =>
    !w ? null : isWorkbookModel(w) ? (w as WorkbookModel) : jsonToWorkbook(w as JsonInput, props.jsonOptions)
  const [activeSheet, setActiveSheet] = useState(0)
  const [zoom, setZoom] = useState(1)
  // 运行时模板(P3 进阶):工具栏导入会覆盖 props.templateFile
  const [runtimeTemplateSrc, setRuntimeTemplateSrc] = useState<ExcelSource | null>(null)
  const [runtimeTemplateName, setRuntimeTemplateName] = useState<string | null>(null)
  const effectiveTemplateSrc = runtimeTemplateSrc ?? props.templateFile ?? null
  const effectiveTemplateName = runtimeTemplateName ?? props.templateName ?? ''
  const clearRuntimeTemplate = () => { setRuntimeTemplateSrc(null); setRuntimeTemplateName(null) }
  const templateInputRef = useRef<HTMLInputElement | null>(null)
  const openTemplateFilePicker = () => templateInputRef.current?.click()
  const onTemplateFilePicked: React.ChangeEventHandler<HTMLInputElement> = (e) => {
    const f = e.target.files?.[0]
    if (!f) return
    setRuntimeTemplateSrc(f)
    setRuntimeTemplateName(f.name)
    e.target.value = ''
  }
  const displayFileName = props.fileName || (props.workbook ? 'JSON 数据' : workbook?.sheets[0]?.name || '')
  const [findOpen, setFindOpen] = useState(false)
  const [, force] = useReducer((x: number) => x + 1, 0)
  // 公式栏编辑态(draft = 编辑中的文本;ref 标记是否正在编辑栏,改 ref 不触发重渲)
  const [fbDraft, setFbDraft] = useState('')
  const fbEditingRef = useRef(false)

  // DOM refs
  const renderAreaRef = useRef<HTMLDivElement>(null)
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const scrollerRef = useRef<HTMLDivElement>(null)
  const spacerRef = useRef<HTMLDivElement>(null)
  const ovMain = useRef<HTMLDivElement>(null)
  const ovFRow = useRef<HTMLDivElement>(null)
  const ovFCol = useRef<HTMLDivElement>(null)
  const ovCorner = useRef<HTMLDivElement>(null)

  const controllerRef = useRef<ViewerController | null>(null)
  const tooltipRef = useRef<TooltipState | null>(null)
  const pluginOvRef = useRef<HTMLDivElement>(null)
  const editorSlotRef = useRef<HTMLDivElement>(null)
  const pluginHostRef = useRef<PluginOverlayHost | null>(null)
  const pluginHandlersRef = useRef<Map<PluginEvent, Set<(p: unknown) => void>>>(new Map())
  // 最新 props / 派生量(供 mount 时注册的 hook 读到当前值)
  const propsRef = useRef(props)
  propsRef.current = props
  const plugins = props.plugins ?? []
  const pluginsRef = useRef(plugins)
  pluginsRef.current = plugins
  const workbookRef = useRef(workbook)
  workbookRef.current = workbook
  const activeSheetRef = useRef(activeSheet)
  activeSheetRef.current = activeSheet

  // 插件 + props 合并:主题 / cellStyle / transformModel(与 Vue 壳同构)
  function buildRendererOpts() {
    const ps = pluginsRef.current
    const theme = Object.assign({}, ...ps.map((p) => p.theme || {}), propsRef.current.theme || {})
    const fns: CellStyleFn[] = ps.map((p) => p.cellStyle).filter(Boolean) as CellStyleFn[]
    if (propsRef.current.cellStyle) fns.push(propsRef.current.cellStyle)
    const cellStyle: CellStyleFn | undefined = fns.length
      ? (cell, pos) => {
          let acc: ReturnType<CellStyleFn> | undefined
          for (const fn of fns) {
            const o = fn(cell, pos)
            if (o) acc = { ...(acc || {}), ...o }
          }
          return acc
        }
      : undefined
    return { theme, cellStyle, cellImageFit: propsRef.current.cellImageFit }
  }
  function effectiveTransform(wb: WorkbookModel): WorkbookModel {
    let m = wb
    for (const p of pluginsRef.current) if (p.transformModel) m = p.transformModel(m) ?? m
    if (propsRef.current.transformModel) m = propsRef.current.transformModel(m) ?? m
    return m
  }
  function buildEditConfig(): EditConfig {
    const p = propsRef.current
    return {
      editable: p.editable,
      cellReadOnly: p.cellReadOnly,
      readOnlyRanges: p.readOnlyRanges,
      editableTargets: p.editableTargets,
      recalc: p.recalc,
      formulaEngine: p.formulaEngine,
    }
  }
  // E2: 合并编辑器解析器(prop 优先,其次插件数组序首个非空)。无任何 editor → undefined
  function editorResolver(): EditorResolver | undefined {
    const hasAny = !!propsRef.current.editor || pluginsRef.current.some((p) => p.editor)
    if (!hasAny) return undefined
    return (cell, pos) => {
      const fromProp = propsRef.current.editor?.(cell, pos)
      if (fromProp) return fromProp as CellEditorFactory
      for (const p of pluginsRef.current) {
        const f = p.editor?.(cell, pos)
        if (f) return f
      }
    }
  }
  /** 派发交互事件给插件(props 回调在各 hook 里另外调) */
  const firePlugin = (event: PluginEvent, payload: unknown) =>
    pluginHandlersRef.current.get(event)?.forEach((h) => h(payload))
  // 导出失败:始终 console.error(没接 onError 也能看到原因,如"请先安装 jspdf")+ 转 onError + alert(与 Vue 壳一致)
  const onExportError = (e: unknown) => {
    const msg = String((e as Error)?.message ?? e)
    console.error('[ooxml-excel-editor] 导出失败:', e)
    propsRef.current.onError?.(msg)
    if (typeof window !== 'undefined' && window.alert) window.alert(msg)
  }

  // 内置导出进度遮罩(P1.5):状态 + 内置 AbortController + onProgress 链接
  const [exportState, setExportState] = useState<ExportProgress | null>(null)
  const [exportBusy, setExportBusy] = useState(false)
  const exportCtrlRef = useRef<AbortController | null>(null)
  const cancelExport = () => exportCtrlRef.current?.abort()

  function chain<T, O extends { onProgress?: (p: ExportProgress) => void; signal?: AbortSignal } | undefined>(
    userOpts: O,
    run: (opts: O) => Promise<T>,
  ): Promise<T> {
    if (propsRef.current.exportProgress === false) return run(userOpts)
    const ctrl = new AbortController()
    exportCtrlRef.current = ctrl
    if (userOpts?.signal) {
      if (userOpts.signal.aborted) ctrl.abort()
      else userOpts.signal.addEventListener('abort', () => ctrl.abort(), { once: true })
    }
    setExportBusy(true)
    setExportState(null)
    const onProgress = (p: ExportProgress) => {
      setExportState(p)
      userOpts?.onProgress?.(p)
    }
    const merged = { ...(userOpts ?? {}), onProgress, signal: ctrl.signal } as O
    return run(merged).finally(() => {
      setExportBusy(false)
      setExportState(null)
      exportCtrlRef.current = null
    })
  }
  function renderPluginOverlays() {
    const host = pluginHostRef.current
    const controller = controllerRef.current
    if (!host || !controller) return
    const ctx: OverlayContext = {
      rectOf: (r, c) => controller.rectOf(r, c),
      rectOfRange: (r) => controller.rectOfRange(r),
      tick: 0,
      workbook: workbookRef.current,
    }
    host.render(pluginsRef.current, ctx)
  }

  // ---- 实例化控制器(一次)。用 layout effect: 在 paint 前就绪,后续 rebuild 也是 layout,顺序确定。 ----
  useLayoutEffect(() => {
    const c = canvasRef.current
    const ra = renderAreaRef.current
    const sc = scrollerRef.current
    const sp = spacerRef.current
    if (!c || !ra || !sc || !sp || !editorSlotRef.current || !ovMain.current || !ovFRow.current || !ovFCol.current || !ovCorner.current) return
    const controller = new ViewerController(
      {
        canvas: c,
        renderArea: ra,
        scroller: sc,
        spacer: sp,
        overlays: { main: ovMain.current, frow: ovFRow.current, fcol: ovFCol.current, corner: ovCorner.current },
        editorSlot: editorSlotRef.current,
      },
      {
        onRenderer: () => force(),
        onRenderTick: () => renderPluginOverlays(), // 插件 overlay 随每帧重定位(纯 DOM,不触发 React 重渲)
        onSelectionChange: () => force(),
        onCellClick: (row, col, text) => {
          const p = { row, col, text }
          propsRef.current.onCellClick?.(p)
          firePlugin('cell-click', p)
        },
        onCellDblClick: (row, col, text) => {
          const p = { row, col, text }
          propsRef.current.onCellDblClick?.(p)
          firePlugin('cell-dblclick', p)
        },
        onHyperlink: (url, cell) => {
          propsRef.current.onHyperlinkClick?.({ url, cell })
          firePlugin('hyperlink-click', { url, cell })
          if (propsRef.current.openLinks !== false) window.open(url, '_blank', 'noopener')
        },
        onTooltip: (tip) => {
          tooltipRef.current = tip
          force()
        },
        onFindChange: () => force(),
        onFilterChange: () => force(),
        onEditEvent: (event, payload) => {
          if (event === 'cell-change') {
            force() // 编辑改了格内容 → 重渲 chrome(公式栏联动);cell-change 非高频,安全
            propsRef.current.onCellChange?.(payload as CellChangePayload)
          }
          else if (event === 'edit-start') propsRef.current.onEditStart?.(payload)
          else if (event === 'edit-commit') propsRef.current.onEditCommit?.(payload)
          else if (event === 'dim-change') propsRef.current.onDimChange?.(payload as DimChangePayload)
          else if (event === 'dirty-change') propsRef.current.onDirtyChange?.(payload as DirtyChangePayload)
          else if (event === 'image-change') propsRef.current.onImageChange?.(payload as ImageChangePayload)
          else if (event === 'struct-change') propsRef.current.onStructChange?.(payload as StructChangePayload)
          firePlugin(event, payload)
        },
        onContextMenuBefore: (payload) => {
          // 1) 先跑插件 contextMenu 串行(后者拿前者的输出)
          for (const p of (propsRef.current.plugins ?? [])) {
            if (p.contextMenu) {
              const next = p.contextMenu(payload.ctx, payload.items)
              if (Array.isArray(next)) payload.items.splice(0, payload.items.length, ...next)
            }
          }
          // 2) 调用方 onBeforeContextMenu
          propsRef.current.onBeforeContextMenu?.(payload)
          // 3) `contextMenu={false}` 直接阻止
          if (propsRef.current.contextMenu === false) payload.preventDefault()
        },
        onContextMenuShow: (payload) => propsRef.current.onContextMenuShow?.(payload),
      },
    )
    controller.fileName = propsRef.current.fileName
    controller.setEditConfig(buildEditConfig())
    controller.setEditorResolver(editorResolver())
    controller.setContextMenuTransform(typeof propsRef.current.contextMenu === 'function' ? propsRef.current.contextMenu : null)
    controller.setLightboxEnabled(propsRef.current.imageLightbox !== false) // 图片点击放大(默认开)
    controllerRef.current = controller
    if (pluginOvRef.current) pluginHostRef.current = new PluginOverlayHost(pluginOvRef.current)

    const ro = new ResizeObserver(() => {
      controller.measure()
      controller.render()
    })
    ro.observe(ra)
    return () => {
      ro.disconnect()
      controller.dispose()
      pluginHostRef.current?.dispose()
      pluginHostRef.current = null
      controllerRef.current = null
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // ---- 载入 src / workbook / templateFile / runtimeTemplate(P3 重设计 2026-06-08) ----
  // 语义同 Vue 壳: xlsx 数据源直接 load + 模板 warn 忽略; JSON 数据源 + 模板 → 解析后 applyStyleTemplate
  useEffect(() => {
    const tplSrc = effectiveTemplateSrc
    const initDataWb = resolveWb(props.workbook)

    async function parseTemplate(src: ExcelSource): Promise<WorkbookModel> {
      const buffer = await loadArrayBuffer(src)
      const fmt = detectFormat(buffer)
      if (fmt === 'xls') throw new Error('模板文件是旧版 .xls 或加密,仅支持 .xlsx/.xlsm')
      if (fmt === 'not-zip') throw new Error('模板文件不是有效的 .xlsx(非 ZIP 包)')
      if (fmt === 'empty') throw new Error('模板文件为空')
      const model = await parseInWorker(buffer)
      finalizeImages(model)
      return model
    }

    async function run() {
      if (props.src) {
        if (tplSrc) console.warn('[ooxml-excel-editor] templateFile 只在 workbook (JSON / 模型) 数据源下生效;xlsx 数据源已自带格式,模板已忽略.')
        await load(props.src, effectiveTransform)
        return
      }
      if (initDataWb && tplSrc) {
        try {
          const tplWb = await parseTemplate(tplSrc)
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
    void run()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.src, props.workbook, props.templateFile, runtimeTemplateSrc])

  // ---- 文件名同步 ----
  useEffect(() => {
    if (controllerRef.current) controllerRef.current.fileName = props.fileName
  }, [props.fileName])

  // ---- 编辑配置同步(E0) ----
  useEffect(() => {
    controllerRef.current?.setEditConfig(buildEditConfig())
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.editable, props.cellReadOnly, props.readOnlyRanges, props.editableTargets, props.recalc, props.formulaEngine])

  // ---- 右键菜单 transform 同步(Plan C) ----
  useEffect(() => {
    controllerRef.current?.setContextMenuTransform(
      typeof props.contextMenu === 'function' ? props.contextMenu : null,
    )
  }, [props.contextMenu])

  // ---- 编辑器解析器同步(E2) ----
  useEffect(() => {
    controllerRef.current?.setEditorResolver(editorResolver())
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.editor, props.plugins])

  // ---- WPS 内嵌图贴合方式同步 ----
  useEffect(() => {
    if (props.cellImageFit) controllerRef.current?.setCellImageFit(props.cellImageFit)
  }, [props.cellImageFit])

  // ---- 图片放大灯箱开关同步 ----
  useEffect(() => {
    controllerRef.current?.setLightboxEnabled(props.imageLightbox !== false)
  }, [props.imageLightbox])

  // ---- 新工作簿 → 选活动表 + onRendered ----
  useEffect(() => {
    if (!workbook) return
    setActiveSheet(workbook.activeSheet)
    propsRef.current.onRendered?.(workbook)
  }, [workbook])

  // ---- 报错回调 ----
  useEffect(() => {
    if (error) propsRef.current.onError?.(error)
  }, [error])

  // ---- 重建渲染器(工作簿 / 活动表 / 主题 / cellStyle / 插件 变化)。layout effect: 同步绘制。----
  useLayoutEffect(() => {
    const controller = controllerRef.current
    if (!controller || !workbook) return
    const sheet: SheetModel | null = workbook.sheets[activeSheet] ?? workbook.sheets[0] ?? null
    if (!sheet) return
    controller.rebuild(sheet, workbook, zoom, buildRendererOpts())
    controller.setSourceBuffer(sourceBuffer) // 注入原件字节(overlay 高保真导出)
    renderPluginOverlays()
    const payload = { index: activeSheet, name: sheet.name }
    propsRef.current.onSheetChange?.(payload)
    firePlugin('sheet-change', payload)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [workbook, activeSheet, props.theme, props.cellStyle, props.plugins, sourceBuffer])

  // ---- 缩放 ----
  useEffect(() => {
    controllerRef.current?.setZoom(zoom)
  }, [zoom])

  // ---- 选区变化 → 回调(读最新选区) ----
  const lastSelKey = useRef('')
  useEffect(() => {
    const controller = controllerRef.current
    if (!controller) return
    const sel = controller.getSelection()
    const active = controller.getActiveCell()
    const key = sel && active ? `${sel.top},${sel.left},${sel.bottom},${sel.right}` : ''
    if (key && key !== lastSelKey.current && active) {
      lastSelKey.current = key
      const payload = { range: sel!, active }
      propsRef.current.onSelectionChange?.(payload)
      firePlugin('selection-change', payload)
    }
  })

  // ---- 命令式句柄 ----
  const dataSheet = (si?: number): SheetModel | null => workbook?.sheets[si ?? activeSheet] ?? null
  const date1904 = workbook?.date1904 ?? false
  const withDate1904 = <T extends ReadOptions>(opts?: T): T => ({ ...(opts as T), date1904 })
  useImperativeHandle(
    ref,
    (): ExcelViewerHandle => ({
      load: (src) => load(src, props.transformModel),
      getWorkbook: () => workbook,
      getActiveSheet: () => activeSheet,
      setActiveSheet: (i) => workbook?.sheets[i] && setActiveSheet(i),
      getSelection: () => controllerRef.current?.getSelection() ?? null,
      setSelection: (range) => controllerRef.current?.setSelectionRange(range),
      rectOf: (row, col) => controllerRef.current?.rectOf(row, col) ?? null,
      rectOfRange: (range) => controllerRef.current?.rectOfRange(range) ?? null,
      redraw: () => controllerRef.current?.render(),
      isCellEditable: (row, col) => controllerRef.current?.isCellEditable(row, col) ?? false,
      setEditableTargets: (targets) => controllerRef.current?.setEditableTargets(targets),
      getEditableTargets: () => controllerRef.current?.getEditableTargets(),
      editCell: (row, col, value) => controllerRef.current?.editCell(row, col, value) ?? false,
      editRange: (range, values) => controllerRef.current?.editRange(range, values) ?? false,
      clearRange: (range) => controllerRef.current?.clearRange(range) ?? false,
      setStyle: (range, patch) => controllerRef.current?.setStyle(range, patch) ?? false,
      getActiveFillColor: () => controllerRef.current?.getActiveFillColor() ?? '#FFFFFF',
      getActiveFontColor: () => controllerRef.current?.getActiveFontColor() ?? '#000000',
      setSelectionFill: (c) => controllerRef.current?.setSelectionFill(c) ?? false,
      setSelectionFontColor: (c) => controllerRef.current?.setSelectionFontColor(c) ?? false,
      getSelectionWrapState: () => controllerRef.current?.getSelectionWrapState() ?? 'none',
      toggleWrapTextOnSelection: () => controllerRef.current?.toggleWrapTextOnSelection() ?? false,
      mergeCells: (range) => controllerRef.current?.mergeCells(range) ?? false,
      unmergeCells: (range) => controllerRef.current?.unmergeCells(range) ?? false,
      pasteText: (text, at) => controllerRef.current?.pasteText(text, at) ?? false,
      pasteRichHtml: (html, at) => controllerRef.current?.pasteRichHtml(html, at) ?? false,
      pasteImageBlob: (blob, at) => controllerRef.current?.pasteImageBlob(blob, at) ?? Promise.resolve(false),
      getImages: () => controllerRef.current?.getImages() ?? [],
      addImage: (a) => controllerRef.current?.addImage(a) ?? -1,
      removeImage: (i) => controllerRef.current?.removeImage(i) ?? false,
      moveImage: (i, dx, dy) => controllerRef.current?.moveImage(i, dx, dy) ?? false,
      resizeImage: (i, w, h) => controllerRef.current?.resizeImage(i, w, h) ?? false,
      getCellImages: () => controllerRef.current?.getCellImages() ?? [],
      getCellImageAt: (row, col) => controllerRef.current?.getCellImageAt(row, col) ?? null,
      openImageLightbox: (src, fileName, mime) => controllerRef.current?.openImageLightbox(src, fileName, mime),
      getCellEditString: () => controllerRef.current?.getCellEditString() ?? '',
      canEditActiveCell: () => controllerRef.current?.canEditActiveCell() ?? false,
      commitActiveCellValue: (v, m) => controllerRef.current?.commitActiveCellValue(v, m) ?? false,
      setCellImageFit: (fit) => controllerRef.current?.setCellImageFit(fit),
      convertImageToCell: (i, row, col) => controllerRef.current?.convertImageToCell(i, row, col) ?? false,
      convertImageToCellAuto: (i) => controllerRef.current?.convertImageToCellAuto(i) ?? false,
      convertAllImagesToCells: (col) => controllerRef.current?.convertAllImagesToCells(col) ?? 0,
      convertImagesInRangeToCell: (range) => chain<number, { onProgress?: (p: ExportProgress) => void; signal?: AbortSignal }>({}, async (o) => { o.onProgress?.({ stage: 'convert', label: '选区浮动图批量嵌入…' }); return controllerRef.current?.convertImagesInRangeToCell(range) ?? 0 }),
      convertCellImagesInRangeToFloat: (range, size) => chain<number, { onProgress?: (p: ExportProgress) => void; signal?: AbortSignal }>({}, async (o) => { o.onProgress?.({ stage: 'convert', label: '选区内嵌图批量浮动化…' }); return controllerRef.current?.convertCellImagesInRangeToFloat(range, size) ?? 0 }),
      openContextMenu: (x, y, items) => controllerRef.current?.openContextMenu(x, y, items),
      closeContextMenu: () => controllerRef.current?.closeContextMenu(),
      convertCellImageToFloat: (row, col, size) => controllerRef.current?.convertCellImageToFloat(row, col, size) ?? false,
      insertRows: (at, count) => controllerRef.current?.insertRows(at, count) ?? false,
      deleteRows: (at, count) => controllerRef.current?.deleteRows(at, count) ?? false,
      insertCols: (at, count) => controllerRef.current?.insertCols(at, count) ?? false,
      deleteCols: (at, count) => controllerRef.current?.deleteCols(at, count) ?? false,
      undo: () => controllerRef.current?.undo(),
      redo: () => controllerRef.current?.redo(),
      canUndo: () => controllerRef.current?.canUndo() ?? false,
      canRedo: () => controllerRef.current?.canRedo() ?? false,
      getEditingCell: () => controllerRef.current?.getEditingCell() ?? null,
      getCellSnapshot: (row, col) => controllerRef.current?.getCellSnapshot(row, col) ?? null,
      inspectCell: (row, col) => controllerRef.current?.inspectCell(row, col) ?? null,
      beginEdit: (row, col) => controllerRef.current?.beginEdit(row, col) ?? false,
      cancelEdit: () => controllerRef.current?.cancelEdit(),
      isEditing: () => controllerRef.current?.isEditing() ?? false,
      setColumnWidth: (col, width) => controllerRef.current?.setColumnWidth(col, width) ?? false,
      setRowHeight: (row, height) => controllerRef.current?.setRowHeight(row, height) ?? false,
      isRecalcReady: () => controllerRef.current?.isRecalcReady() ?? false,
      getVirtualExtent: () => controllerRef.current?.getVirtualExtent() ?? { rows: 0, cols: 0 },
      isDirty: () => controllerRef.current?.isDirty() ?? false,
      resetToOriginal: () => controllerRef.current?.resetToOriginal() ?? false,
      exportImage: (opts) => chain(opts, (o) => controllerRef.current!.exportImage(o)),
      downloadImage: (opts) => chain(opts, (o) => controllerRef.current!.downloadImage(o)),
      exportPdf: (opts) => chain(opts, (o) => controllerRef.current!.exportPdf(o)),
      downloadPdf: (opts) => chain(opts, (o) => controllerRef.current!.downloadPdf(o)),
      print: (opts) => chain(opts, (o) => controllerRef.current!.print(o)),
      exportXlsx: (opts) => chain(opts, (o) => controllerRef.current!.exportXlsx(o)),
      downloadXlsx: (opts) => chain(opts, (o) => controllerRef.current!.downloadXlsx(o)),
      exportJson: (opts) => controllerRef.current?.exportJson(opts) ?? '{}',
      downloadJson: (opts) => controllerRef.current?.downloadJson(opts),
      exportCsv: (opts) => controllerRef.current?.exportCsv(opts) ?? '',
      downloadCsv: (opts) => controllerRef.current?.downloadCsv(opts),
      getCellValue: (row, col, si) => {
        const s = dataSheet(si)
        return s ? getCellValue(s, row, col) : null
      },
      getCellText: (row, col, si) => {
        const s = dataSheet(si)
        return s ? getCellText(s, row, col, date1904) : ''
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
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [workbook, activeSheet],
  )

  // ---- 给插件 / 工具栏用的稳定命令式 API(读 ref,跨渲染稳定) ----
  const apiSheet = (si?: number): SheetModel | null => workbookRef.current?.sheets[si ?? activeSheetRef.current] ?? null
  const apiDate1904 = () => workbookRef.current?.date1904 ?? false
  const viewerApi = useRef<ViewerApi>({
    load: (src) => load(src, effectiveTransform),
    getWorkbook: () => workbookRef.current,
    getActiveSheet: () => activeSheetRef.current,
    setActiveSheet: (i) => {
      if (workbookRef.current?.sheets[i]) setActiveSheet(i)
    },
    getSelection: () => controllerRef.current?.getSelection() ?? null,
    setSelection: (range) => controllerRef.current?.setSelectionRange(range),
    rectOf: (row, col) => controllerRef.current?.rectOf(row, col) ?? null,
    rectOfRange: (range) => controllerRef.current?.rectOfRange(range) ?? null,
    redraw: () => controllerRef.current?.render(),
    isCellEditable: (row, col) => controllerRef.current?.isCellEditable(row, col) ?? false,
    setEditableTargets: (targets) => controllerRef.current?.setEditableTargets(targets),
    getEditableTargets: () => controllerRef.current?.getEditableTargets(),
    editCell: (row, col, value) => controllerRef.current?.editCell(row, col, value) ?? false,
    editRange: (range, values) => controllerRef.current?.editRange(range, values) ?? false,
    clearRange: (range) => controllerRef.current?.clearRange(range) ?? false,
    setStyle: (range, patch) => controllerRef.current?.setStyle(range, patch) ?? false,
    getActiveFillColor: () => controllerRef.current?.getActiveFillColor() ?? '#FFFFFF',
    getActiveFontColor: () => controllerRef.current?.getActiveFontColor() ?? '#000000',
    setSelectionFill: (c) => controllerRef.current?.setSelectionFill(c) ?? false,
    setSelectionFontColor: (c) => controllerRef.current?.setSelectionFontColor(c) ?? false,
    getSelectionWrapState: () => controllerRef.current?.getSelectionWrapState() ?? 'none',
    toggleWrapTextOnSelection: () => controllerRef.current?.toggleWrapTextOnSelection() ?? false,
    mergeCells: (range) => controllerRef.current?.mergeCells(range) ?? false,
    unmergeCells: (range) => controllerRef.current?.unmergeCells(range) ?? false,
    pasteText: (text, at) => controllerRef.current?.pasteText(text, at) ?? false,
    pasteRichHtml: (html, at) => controllerRef.current?.pasteRichHtml(html, at) ?? false,
    pasteImageBlob: (blob, at) => controllerRef.current?.pasteImageBlob(blob, at) ?? Promise.resolve(false),
    getImages: () => controllerRef.current?.getImages() ?? [],
    addImage: (a) => controllerRef.current?.addImage(a) ?? -1,
    removeImage: (i) => controllerRef.current?.removeImage(i) ?? false,
    moveImage: (i, dx, dy) => controllerRef.current?.moveImage(i, dx, dy) ?? false,
    resizeImage: (i, w, h) => controllerRef.current?.resizeImage(i, w, h) ?? false,
    getCellImages: () => controllerRef.current?.getCellImages() ?? [],
    getCellImageAt: (row, col) => controllerRef.current?.getCellImageAt(row, col) ?? null,
    openImageLightbox: (src, fileName, mime) => controllerRef.current?.openImageLightbox(src, fileName, mime),
    getCellEditString: () => controllerRef.current?.getCellEditString() ?? '',
    canEditActiveCell: () => controllerRef.current?.canEditActiveCell() ?? false,
    commitActiveCellValue: (v, m) => controllerRef.current?.commitActiveCellValue(v, m) ?? false,
    setCellImageFit: (fit) => controllerRef.current?.setCellImageFit(fit),
    convertImageToCell: (i, row, col) => controllerRef.current?.convertImageToCell(i, row, col) ?? false,
    convertImageToCellAuto: (i) => controllerRef.current?.convertImageToCellAuto(i) ?? false,
    convertAllImagesToCells: (col) => controllerRef.current?.convertAllImagesToCells(col) ?? 0,
    convertImagesInRangeToCell: (range) => chain<number, { onProgress?: (p: ExportProgress) => void; signal?: AbortSignal }>({}, async (o) => { o.onProgress?.({ stage: 'convert', label: '选区浮动图批量嵌入…' }); return controllerRef.current?.convertImagesInRangeToCell(range) ?? 0 }),
    convertCellImagesInRangeToFloat: (range, size) => chain<number, { onProgress?: (p: ExportProgress) => void; signal?: AbortSignal }>({}, async (o) => { o.onProgress?.({ stage: 'convert', label: '选区内嵌图批量浮动化…' }); return controllerRef.current?.convertCellImagesInRangeToFloat(range, size) ?? 0 }),
    openContextMenu: (x, y, items) => controllerRef.current?.openContextMenu(x, y, items),
    closeContextMenu: () => controllerRef.current?.closeContextMenu(),
    convertCellImageToFloat: (row, col, size) => controllerRef.current?.convertCellImageToFloat(row, col, size) ?? false,
    insertRows: (at, count) => controllerRef.current?.insertRows(at, count) ?? false,
    deleteRows: (at, count) => controllerRef.current?.deleteRows(at, count) ?? false,
    insertCols: (at, count) => controllerRef.current?.insertCols(at, count) ?? false,
    deleteCols: (at, count) => controllerRef.current?.deleteCols(at, count) ?? false,
    undo: () => controllerRef.current?.undo(),
    redo: () => controllerRef.current?.redo(),
    canUndo: () => controllerRef.current?.canUndo() ?? false,
    canRedo: () => controllerRef.current?.canRedo() ?? false,
    getEditingCell: () => controllerRef.current?.getEditingCell() ?? null,
    getCellSnapshot: (row, col) => controllerRef.current?.getCellSnapshot(row, col) ?? null,
    inspectCell: (row, col) => controllerRef.current?.inspectCell(row, col) ?? null,
    beginEdit: (row, col) => controllerRef.current?.beginEdit(row, col) ?? false,
    cancelEdit: () => controllerRef.current?.cancelEdit(),
    isEditing: () => controllerRef.current?.isEditing() ?? false,
    setColumnWidth: (col, width) => controllerRef.current?.setColumnWidth(col, width) ?? false,
    setRowHeight: (row, height) => controllerRef.current?.setRowHeight(row, height) ?? false,
    isRecalcReady: () => controllerRef.current?.isRecalcReady() ?? false,
    getVirtualExtent: () => controllerRef.current?.getVirtualExtent() ?? { rows: 0, cols: 0 },
    isDirty: () => controllerRef.current?.isDirty() ?? false,
    resetToOriginal: () => controllerRef.current?.resetToOriginal() ?? false,
    exportImage: (opts) => chain(opts, (o) => controllerRef.current!.exportImage(o)),
    downloadImage: (opts) => chain(opts, (o) => controllerRef.current!.downloadImage(o)),
    exportPdf: (opts) => chain(opts, (o) => controllerRef.current!.exportPdf(o)),
    downloadPdf: (opts) => chain(opts, (o) => controllerRef.current!.downloadPdf(o)),
    print: (opts) => chain(opts, (o) => controllerRef.current!.print(o)),
    exportXlsx: (opts) => chain(opts, (o) => controllerRef.current!.exportXlsx(o)),
    downloadXlsx: (opts) => chain(opts, (o) => controllerRef.current!.downloadXlsx(o)),
    exportJson: (opts) => controllerRef.current?.exportJson(opts) ?? '{}',
    downloadJson: (opts) => controllerRef.current?.downloadJson(opts),
    exportCsv: (opts) => controllerRef.current?.exportCsv(opts) ?? '',
    downloadCsv: (opts) => controllerRef.current?.downloadCsv(opts),
    getCellValue: (row, col, si) => {
      const s = apiSheet(si)
      return s ? getCellValue(s, row, col) : null
    },
    getCellText: (row, col, si) => {
      const s = apiSheet(si)
      return s ? getCellText(s, row, col, apiDate1904()) : ''
    },
    getSheetData: (opts, si) => {
      const s = apiSheet(si)
      return s ? getSheetData(s, { ...opts, date1904: apiDate1904() }) : []
    },
    getSheetJSON: (opts, si) => {
      const s = apiSheet(si)
      return s ? sheetToJSON(s, { ...opts, date1904: apiDate1904() }) : []
    },
    getRangeData: (range, opts, si) => {
      const s = apiSheet(si)
      return s ? getRangeData(s, range, { ...opts, date1904: apiDate1904() }) : []
    },
  }).current

  // ---- 插件 setup / events 注册(plugins 变化时重建) ----
  useEffect(() => {
    const handlers = new Map<PluginEvent, Set<(p: unknown) => void>>()
    const cleanups: Array<() => void> = []
    const register = (event: PluginEvent, fn: (p: unknown) => void) => {
      let set = handlers.get(event)
      if (!set) handlers.set(event, (set = new Set()))
      set.add(fn)
    }
    for (const p of plugins) {
      if (p.events)
        for (const [ev, fn] of Object.entries(p.events)) if (fn) register(ev as PluginEvent, fn as (x: unknown) => void)
      const cleanup = p.setup?.({ viewer: viewerApi, on: register, redraw: () => controllerRef.current?.render() })
      if (typeof cleanup === 'function') cleanups.push(cleanup)
    }
    pluginHandlersRef.current = handlers
    renderPluginOverlays()
    return () => {
      cleanups.forEach((fn) => fn())
      pluginHandlersRef.current = new Map()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.plugins])

  // ---- 根容器 Ctrl/Cmd+F 打开查找 ----
  const onRootKeyDown = (e: React.KeyboardEvent) => {
    if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) {
      e.preventDefault()
      setFindOpen(true)
    }
  }
  const closeFind = () => {
    setFindOpen(false)
    controllerRef.current?.clearFind()
    scrollerRef.current?.focus()
  }

  const onScroll = () => {
    const sc = scrollerRef.current
    if (!sc) return
    tooltipRef.current = null
    controllerRef.current?.setScroll(sc.scrollLeft, sc.scrollTop)
  }

  // ---- 渲染期读 controller 派生状态(chrome) ----
  const controller = controllerRef.current
  const renderer = controller?.renderer ?? null
  const active = controller?.getActiveCell() ?? null
  const selection = controller?.getSelection() ?? null
  const activeAddr = active ? colIndexToLetters(active.col) + (active.row + 1) : ''
  const formulaText = renderer && active ? (renderer.cellFormula(active.row, active.col) ?? renderer.cellText(active.row, active.col)) : ''
  // 公式栏:可编辑判定 + 编辑字符串 + 联动处理(提交→改格;切格/格内编辑→栏更新)
  const fbCanEdit = !!controller?.canEditActiveCell()
  const fbEditString = controller?.getCellEditString() ?? ''
  const fbValue = fbEditingRef.current ? fbDraft : fbEditString
  const fbFocus = () => {
    fbEditingRef.current = true
    setFbDraft(controller?.getCellEditString() ?? '')
  }
  const fbCommit = (move?: 'down') => {
    controller?.commitActiveCellValue(fbValue, move)
    fbEditingRef.current = false
    if (move === 'down') scrollerRef.current?.focus()
    force()
  }
  const fbCancel = () => {
    fbEditingRef.current = false
    scrollerRef.current?.focus()
    force()
  }
  const fbBlur = () => {
    if (fbEditingRef.current) fbCommit()
  }
  const fbKeydown = (e: React.KeyboardEvent) => {
    e.stopPropagation() // 别让网格键盘处理插手
    if (e.key === 'Enter') {
      e.preventDefault()
      fbCommit('down')
    } else if (e.key === 'Escape') {
      e.preventDefault()
      fbCancel()
    }
  }
  const rangeLabel =
    selection && !(selection.top === selection.bottom && selection.left === selection.right)
      ? `${colIndexToLetters(selection.left)}${selection.top + 1}:${colIndexToLetters(selection.right)}${selection.bottom + 1}`
      : ''
  const stats = renderer && selection ? renderer.selectionStats(selection) : null
  const findState = controller?.getFindState() ?? { query: '', matchCase: false, wholeCell: false, count: 0, index: -1 }
  const filterPopup = controller?.getFilterPopup() ?? null
  const tooltip = tooltipRef.current
  const curSheet = workbook?.sheets[activeSheet] ?? null
  const visibleSheets = workbook ? workbook.sheets.map((s, i) => ({ s, i })).filter(({ s }) => s.state === 'visible') : []

  return (
    <div className={'rxl' + (props.className ? ' ' + props.className : '')} style={props.style} onKeyDown={onRootKeyDown}>
      {workbook && (displayFileName || effectiveTemplateName) && (
        <div className="rxl-title" title={displayFileName + (effectiveTemplateName ? ' · 模板: ' + effectiveTemplateName : '')}>
          <span className="name">{displayFileName || '未命名工作簿'}</span>
          {effectiveTemplateName && <span className="tpl"> · 模板: {effectiveTemplateName}</span>}
        </div>
      )}
      {/* 工具栏「模板」项的隐藏文件拾取器(P3 进阶) */}
      <input ref={templateInputRef} type="file" accept=".xlsx,.xlsm" hidden onChange={onTemplateFilePicked} />
      {workbook && (
        <div className="rxl-toolbar">
          <button
            className={findOpen ? 'active' : ''}
            onClick={() => (findOpen ? closeFind() : setFindOpen(true))}
            title="查找 (Ctrl+F)"
          >
            查找
          </button>
          <button
            className={curSheet?.autoFilterRange ? 'active' : ''}
            onClick={() => controllerRef.current?.toggleAutoFilter()}
            title="切换自动筛选"
          >
            筛选
          </button>
          <button disabled={!controller?.hasFilters()} onClick={() => controllerRef.current?.clearAllFilters()}>
            清除筛选
          </button>
          <button disabled={!selection} onClick={() => void controllerRef.current?.copySelection()} title="复制 (Ctrl+C)">
            复制
          </button>
          <button onClick={() => void controllerRef.current?.downloadImage().catch(onExportError)}>导出 PNG</button>
          <button onClick={() => void controllerRef.current?.downloadPdf().catch(onExportError)}>导出 PDF</button>
          <select value={Math.round(zoom * 100)} onChange={(e) => setZoom(Number(e.target.value) / 100)} title="缩放">
            {[50, 75, 100, 125, 150, 200].map((p) => (
              <option key={p} value={p}>
                {p}%
              </option>
            ))}
          </select>
          {/* 模板(P3 重设计 2026-06-08):仅在 JSON / 模型数据源下生效;xlsx 数据源禁用 */}
          {(() => {
            const isXlsxSrc = !!props.src && !props.workbook
            return (
              <>
                <button
                  className={effectiveTemplateSrc ? 'active' : ''}
                  disabled={isXlsxSrc}
                  onClick={openTemplateFilePicker}
                  title={isXlsxSrc
                    ? '模板仅对 JSON / 模型数据源生效;当前是 xlsx 数据源,模板不可用'
                    : effectiveTemplateSrc
                      ? `模板已加载: ${effectiveTemplateName || '(未命名)'} — 点击重新导入`
                      : '为 JSON / 模型数据源套用 .xlsx 模板的样式(模板的文字内容会被丢弃)'}
                >
                  模板{effectiveTemplateSrc ? ' ▾' : ''}
                </button>
                {effectiveTemplateSrc && !isXlsxSrc && (
                  <button onClick={clearRuntimeTemplate} title="清除模板,切回默认渲染">
                    清除模板
                  </button>
                )}
              </>
            )
          })()}
          {/* 插件贡献的工具栏按钮(跨框架同一份插件) */}
          {plugins
            .flatMap((p) => p.toolbar ?? [])
            .filter((it) => it.type !== 'separator')
            .map((it) => (
              <button
                key={it.id}
                className={it.active?.(viewerApi) ? 'active' : ''}
                disabled={it.disabled?.(viewerApi)}
                title={it.title}
                onClick={() => it.onClick?.(viewerApi)}
              >
                {it.label ?? it.icon ?? it.id}
              </button>
            ))}
        </div>
      )}

      {workbook && (
        <div className="rxl-formula-bar">
          <span className="addr">{activeAddr || '—'}</span>
          <span className="fx">fx</span>
          {fbCanEdit ? (
            <input
              className="content content-input"
              value={fbValue}
              title={fbValue}
              spellCheck={false}
              onFocus={fbFocus}
              onChange={(e) => setFbDraft(e.target.value)}
              onKeyDown={fbKeydown}
              onBlur={fbBlur}
            />
          ) : (
            <span className="content" title={formulaText}>
              {formulaText}
            </span>
          )}
        </div>
      )}

      <div className="rxl-render-area" ref={renderAreaRef}>
        <canvas ref={canvasRef} className="rxl-canvas" />
        <div className="rxl-ov" ref={ovMain} />
        <div className="rxl-ov" ref={ovFCol} />
        <div className="rxl-ov" ref={ovFRow} />
        <div className="rxl-ov" ref={ovCorner} />
        <div
          className="rxl-scroller"
          ref={scrollerRef}
          tabIndex={0}
          onScroll={onScroll}
          onMouseDown={(e) => controllerRef.current?.onMouseDown(e.nativeEvent)}
          onMouseMove={(e) => controllerRef.current?.onMouseMove(e.nativeEvent)}
          onMouseUp={(e) => controllerRef.current?.onMouseUp(e.nativeEvent)}
          onMouseLeave={() => controllerRef.current?.onMouseLeave()}
          onDoubleClick={(e) => controllerRef.current?.onDblClick(e.nativeEvent)}
          onKeyDown={(e) => controllerRef.current?.onKeyDown(e.nativeEvent)}
          onContextMenu={(e) => controllerRef.current?.onContextMenu(e.nativeEvent)}
        >
          <div className="rxl-spacer" ref={spacerRef} />
        </div>

        {/* 插件 overlay:框架无关 DOM,由 PluginOverlayHost 挂载(随 tick 跟随滚动/缩放) */}
        <div className="rxl-ov-slot">
          <div ref={pluginOvRef} />
        </div>
        {/* 单元格编辑器层(E2):CellEditorHost 挂载 */}
        <div className="rxl-editor-slot" ref={editorSlotRef} />

        {findOpen && workbook && (
          <div className="rxl-findbar">
            <input
              autoFocus
              placeholder="查找…"
              value={findState.query}
              onChange={(e) => controllerRef.current?.setFindQuery(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') (e.shiftKey ? controllerRef.current?.findPrev() : controllerRef.current?.findNext())
                else if (e.key === 'Escape') closeFind()
              }}
            />
            <span className="count">{findState.count ? `${findState.index + 1}/${findState.count}` : '无结果'}</span>
            <button onClick={() => controllerRef.current?.findPrev()}>↑</button>
            <button onClick={() => controllerRef.current?.findNext()}>↓</button>
            <button onClick={closeFind}>✕</button>
          </div>
        )}

        {filterPopup && (
          <FilterPopup
            key={filterPopup.col}
            values={filterPopup.values}
            selected={filterPopup.selected}
            x={filterPopup.x}
            y={filterPopup.y}
            sortDir={filterPopup.sortDir}
            onApply={(checked) => controllerRef.current?.applyFilterSelection(checked)}
            onClear={() => controllerRef.current?.clearFilterColumn()}
            onClose={() => controllerRef.current?.closeFilterPopup()}
            onSort={(dir) => {
              const c = filterPopup.col
              controllerRef.current?.closeFilterPopup()
              controllerRef.current?.sortColumn(c, dir)
            }}
          />
        )}

        {tooltip && (
          <div className={'rxl-tooltip ' + tooltip.kind} style={{ left: tooltip.x, top: tooltip.y }}>
            {tooltip.text}
          </div>
        )}

        {loading && (
          <div className="rxl-state">
            {progress?.stage === 'read' ? '读取文件…' : progress?.stage === 'parse' ? '解析中…' : '构建表格…'}
          </div>
        )}
        {!loading && error && <div className="rxl-state error">解析失败：{error}</div>}
        {!loading && !workbook && <div className="rxl-state">拖入或选择一个 .xlsx 文件</div>}
      </div>

      {workbook && (
        <div className="rxl-status-bar">
          <span style={{ color: '#888' }}>{rangeLabel || activeAddr}</span>
          <div className="grow" />
          {stats && stats.numCount > 0 ? (
            <>
              <span>计数 {stats.count}</span>
              <span>求和 {fmtNum(stats.sum)}</span>
              <span>平均 {fmtNum(stats.avg)}</span>
              <span>最大 {fmtNum(stats.max)}</span>
              <span>最小 {fmtNum(stats.min)}</span>
            </>
          ) : stats && stats.count > 0 ? (
            <span>计数 {stats.count}</span>
          ) : null}
        </div>
      )}

      {workbook && visibleSheets.length > 0 && (
        <div className="rxl-tabs">
          {visibleSheets.map(({ s, i }) => (
            <button key={i} className={i === activeSheet ? 'active' : ''} onClick={() => setActiveSheet(i)}>
              {s.name}
            </button>
          ))}
        </div>
      )}
      {/* 内置导出进度遮罩(P1.5):exportProgress=false 关闭;renderExportProgress 自渲染覆盖 */}
      {props.exportProgress !== false && (
        props.renderExportProgress
          ? props.renderExportProgress({ state: exportState, busy: exportBusy, cancel: cancelExport })
          : <ExportProgressOverlay state={exportState} busy={exportBusy} onCancel={cancelExport} />
      )}
    </div>
  )
})

/** 极简筛选浮层(React 版),与 core 的 FilterPopupState 对齐 */
function FilterPopup(props: {
  values: string[]
  selected: string[]
  x: number
  y: number
  sortDir: 'asc' | 'desc' | null
  onApply: (checked: string[]) => void
  onClear: () => void
  onClose: () => void
  onSort: (dir: 'asc' | 'desc') => void
}) {
  // selected 为空 = 该列未筛选 = 全选
  const initial = props.selected.length ? new Set(props.selected) : new Set(props.values)
  const [checked, setChecked] = useState<Set<string>>(initial)
  const toggle = (v: string) => {
    const next = new Set(checked)
    if (next.has(v)) next.delete(v)
    else next.add(v)
    setChecked(next)
  }
  return (
    <div className="rxl-filterpop" style={{ left: props.x, top: props.y }}>
      <div className="sort">
        <button className={props.sortDir === 'asc' ? 'on' : ''} onClick={() => props.onSort('asc')}>
          ↑ 升序
        </button>
        <button className={props.sortDir === 'desc' ? 'on' : ''} onClick={() => props.onSort('desc')}>
          ↓ 降序
        </button>
      </div>
      <div className="list">
        <label>
          <input
            type="checkbox"
            checked={checked.size === props.values.length}
            ref={(el) => {
              if (el) el.indeterminate = checked.size > 0 && checked.size < props.values.length
            }}
            onChange={(e) => setChecked(e.target.checked ? new Set(props.values) : new Set())}
          />
          (全选)
        </label>
        {props.values.map((v) => (
          <label key={v}>
            <input type="checkbox" checked={checked.has(v)} onChange={() => toggle(v)} />
            {v}
          </label>
        ))}
      </div>
      <div className="foot">
        <button onClick={props.onClear}>清除</button>
        <button onClick={() => props.onApply([...checked])}>确定</button>
        <button onClick={props.onClose}>取消</button>
      </div>
    </div>
  )
}
