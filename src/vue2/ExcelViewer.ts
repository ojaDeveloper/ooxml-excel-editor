/**
 * Vue 2 壳 — render function 实现, 不依赖 SFC 编译器.
 *
 * 为什么 .ts 而非 .vue?
 *   - 项目同时存在 vue@3 (主) 跟 vue@2 (alias) → @vitejs/plugin-vue2 在解析 @vue/compiler-sfc
 *     时拿到 vue@3 的版本, 导致 SFC 编译失败 ("currentInput.slice is not a function").
 *   - 改用 render function: 完全不走 SFC 编译路径, 直接用 vue 2.7 内置 Composition API +
 *     `h()` 渲染. 跟 React 壳 (.tsx + hook) 的写法接近, 工程上更稳.
 *
 * Vue 2.7+ 内置 Composition API (ref/computed/watch/onMounted/onBeforeUnmount/defineComponent),
 * `vue` 通过 vite alias 重定向到 node_modules/vue2 (npm:vue@^2.7.16).
 *
 * MVP 范围 (1.3.0 第一波): 加载 + 渲染 + 基础事件 + 命令式 API. 不带内置工具栏 / 公式栏 / 查找 /
 * 筛选 / 标签 / 导出对话框 / 右键菜单 — 后续迭代补.
 */
// 开发期: 'vue2' alias = node_modules/vue2 (npm:vue@^2.7.16). build 时 rollup output paths
// 把 'vue2' 替换为 'vue' (用户那 vue 2.7+ peer)
import Vue, { defineComponent, h, ref, watch, onMounted, onBeforeUnmount, getCurrentInstance, type PropType } from 'vue2'
import type { ExcelPlugin, ViewerApi } from '@/core/plugin'
import type { ExcelSource } from '@/core/loader'
import type { WorkbookModel, MergeRange, TransformModelFn } from '@/core/model/types'
import { ViewerController } from '@/core/viewer/controller'
import { useExcelDocumentVue2 } from './use-excel-document'
import './excel-viewer.css'

export default defineComponent({
  name: 'OoxmlExcelViewer',
  props: {
    src: { type: [String, Object, Blob, ArrayBuffer, Uint8Array] as PropType<ExcelSource>, default: undefined },
    fileName: { type: String, default: '' },
    editable: { type: Boolean, default: false },
    plugins: { type: Array as PropType<ExcelPlugin[]>, default: () => [] },
    transformModel: { type: Function as PropType<TransformModelFn>, default: undefined },
    openLinks: { type: Boolean, default: true },
  },
  setup(props, { emit, expose }) {
    // Vue 2 render function 里 ref:'xxx' 不会自动连接 setup 的 ref() 变量, 必须通过
    // instance.$refs 拿. (这点跟 Vue 3 不一样)
    const instance = getCurrentInstance()
    const refs = () => instance?.proxy?.$refs as Record<string, HTMLElement | undefined> | undefined

    const { loading: _loading, error, workbook, load, loadModel: _loadModel, sourceBuffer } = useExcelDocumentVue2()
    const activeSheet = ref(0)
    const zoom = ref(1)

    let controller: ViewerController | null = null
    let resizeObserver: ResizeObserver | null = null

    function effectiveTransform(wb: WorkbookModel): WorkbookModel {
      let m = wb
      for (const p of props.plugins) if (p.transformModel) m = p.transformModel(m) ?? m
      if (props.transformModel) m = props.transformModel(m) ?? m
      return m
    }

    onMounted(() => {
      const r = refs()
      if (!r) return
      const canvas = r.canvasEl as HTMLCanvasElement | undefined
      const renderArea = r.renderAreaEl
      const scroller = r.scrollerEl
      const spacer = r.spacerEl
      const editorSlot = r.editorSlotEl
      const ovMain = r.ovMainEl
      const ovFRow = r.ovFRowEl
      const ovFCol = r.ovFColEl
      const ovCorner = r.ovCornerEl
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
          onRenderTick: () => {},
          onSelectionChange: () => {
            const sel = controller?.getSelection()
            const active = controller?.getActiveCell()
            if (sel && active) emit('selection-change', { range: sel, active })
          },
          onCellClick: (row, col, text) => emit('cell-click', { row, col, text }),
          onCellDblClick: (row, col, text) => emit('cell-dblclick', { row, col, text }),
          onHyperlink: (url, cell) => {
            emit('hyperlink-click', { url, cell })
            if (props.openLinks) window.open(url, '_blank', 'noopener')
          },
          onTooltip: () => {},
          onFindChange: () => {},
          onFilterChange: () => {},
          onEditEvent: (event, payload) => {
            if (event === 'cell-change') emit('cell-change', payload)
          },
        },
      )
      controller.fileName = props.fileName
      controller.setEditConfig({ editable: props.editable })
      if (props.src) void load(props.src, effectiveTransform)

      resizeObserver = new ResizeObserver(() => {
        controller?.measure()
        controller?.render()
      })
      resizeObserver.observe(renderArea)
    })

    watch(() => props.src, (src) => {
      if (src) void load(src, effectiveTransform)
    })

    watch(() => props.editable, (v) => controller?.setEditConfig({ editable: v }))

    watch(workbook, async (wb) => {
      if (!wb || !controller) return
      activeSheet.value = wb.activeSheet
      await Vue.nextTick()
      const sheet = wb.sheets[activeSheet.value] ?? wb.sheets[0]
      if (!sheet) return
      controller.rebuild(sheet, wb, zoom.value, {})
      controller.setSourceBuffer(sourceBuffer.value)
      emit('rendered', wb)
    })

    watch(error, (msg) => { if (msg) emit('error', msg) })

    onBeforeUnmount(() => {
      resizeObserver?.disconnect()
      controller?.dispose()
    })

    function onScroll(e: Event) {
      const sc = e.target as HTMLElement
      controller?.setScroll(sc.scrollLeft, sc.scrollTop)
    }

    // 命令式 API (跟 React Handle / Vue 3 viewerApi 对齐, MVP 范围)
    const viewerApi = {
      load: (src: ExcelSource) => load(src, effectiveTransform),
      getWorkbook: () => workbook.value,
      getActiveSheet: () => activeSheet.value,
      setActiveSheet: (i: number) => { if (workbook.value?.sheets[i]) activeSheet.value = i },
      getSelection: () => controller?.getSelection() ?? null,
      setSelection: (range: MergeRange) => controller?.setSelectionRange(range),
      rectOf: (row: number, col: number) => controller?.rectOf(row, col) ?? null,
      rectOfRange: (range: MergeRange) => controller?.rectOfRange(range) ?? null,
      redraw: () => controller?.render(),
      isCellEditable: (row: number, col: number) => controller?.isCellEditable(row, col) ?? false,
      exportImage: (opts?: Parameters<ViewerApi['exportImage']>[0]) => controller!.exportImage(opts),
      downloadImage: (opts?: Parameters<ViewerApi['downloadImage']>[0]) => controller!.downloadImage(opts),
      exportPdf: (opts?: Parameters<ViewerApi['exportPdf']>[0]) => controller!.exportPdf(opts),
      downloadPdf: (opts?: Parameters<ViewerApi['downloadPdf']>[0]) => controller!.downloadPdf(opts),
      exportXlsx: (opts?: Parameters<ViewerApi['exportXlsx']>[0]) => controller!.exportXlsx(opts),
      downloadXlsx: (opts?: Parameters<ViewerApi['downloadXlsx']>[0]) => controller!.downloadXlsx(opts),
      exportCsv: (opts?: Parameters<ViewerApi['exportCsv']>[0]) => controller?.exportCsv(opts) ?? '',
      downloadCsv: (opts?: Parameters<ViewerApi['downloadCsv']>[0]) => controller?.downloadCsv(opts),
      undo: () => controller?.undo(),
      redo: () => controller?.redo(),
      canUndo: () => controller?.canUndo() ?? false,
      canRedo: () => controller?.canRedo() ?? false,
      editCell: (row: number, col: number, value: Parameters<ViewerApi['editCell']>[2]) => controller?.editCell(row, col, value) ?? false,
      getCellValue: (row: number, col: number) => {
        const wb = workbook.value
        const sheet = wb?.sheets[activeSheet.value]
        return sheet?.cells.get(`${row}:${col}`)?.raw ?? null
      },
      getCellText: (row: number, col: number) => controller?.renderer?.cellText(row, col) ?? '',
    }
    expose?.(viewerApi)

    return () => h('div', { ref: 'rootEl', class: 'ooxml-excel-viewer' }, [
      h('div', { ref: 'renderAreaEl', class: 'ov-render-area' }, [
        h('canvas', { ref: 'canvasEl', class: 'ov-grid-canvas' }),
        h('div', { ref: 'ovMainEl', class: 'ov ov-main' }),
        h('div', { ref: 'ovFColEl', class: 'ov ov-fcol' }),
        h('div', { ref: 'ovFRowEl', class: 'ov ov-frow' }),
        h('div', { ref: 'ovCornerEl', class: 'ov ov-corner' }),
        h('div', { ref: 'pluginOvEl', class: 'ov-plugin' }),
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
      ]),
    ])
  },
})
