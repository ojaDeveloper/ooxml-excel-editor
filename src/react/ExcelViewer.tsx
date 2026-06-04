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
import type { CellStyleFn, MergeRange, SheetModel, TransformModelFn, WorkbookModel } from '@/core/model/types'
import type { ViewerTheme } from '@/core/render/theme'
import type { ExcelSource } from '@/core/loader'
import type { ImageExportOptions, PdfExportOptions, PrintOptions } from '@/core/export'
import {
  getCellValue,
  getCellText,
  getSheetData,
  getRangeData,
  sheetToJSON,
  type ReadOptions,
} from '@/core/model/data-access'
import { colIndexToLetters } from '@/core/layout/grid-metrics'
import { ViewerController, type Cell, type TooltipState } from '@/core/viewer/controller'
import { useExcelDocument } from './use-excel-document'
import './excel-viewer.css'

export interface ExcelViewerProps {
  src?: ExcelSource
  fileName?: string
  theme?: Partial<ViewerTheme>
  /** 单击超链接是否自动打开(默认 true) */
  openLinks?: boolean
  transformModel?: TransformModelFn
  cellStyle?: CellStyleFn
  className?: string
  style?: CSSProperties
  onRendered?: (wb: WorkbookModel) => void
  onError?: (msg: string) => void
  onCellClick?: (p: { row: number; col: number; text: string }) => void
  onCellDblClick?: (p: { row: number; col: number; text: string }) => void
  onSelectionChange?: (p: { range: MergeRange; active: Cell }) => void
  onHyperlinkClick?: (p: { url: string; cell: Cell }) => void
  onSheetChange?: (p: { index: number; name: string }) => void
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
  exportImage: (opts?: ImageExportOptions) => Promise<Blob>
  downloadImage: (opts?: ImageExportOptions) => Promise<void>
  exportPdf: (opts?: PdfExportOptions) => Promise<Blob>
  downloadPdf: (opts?: PdfExportOptions) => Promise<void>
  print: (opts?: PrintOptions) => Promise<void>
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
  const { loading, error, workbook, progress, load } = useExcelDocument()
  const [activeSheet, setActiveSheet] = useState(0)
  const [zoom, setZoom] = useState(1)
  const [findOpen, setFindOpen] = useState(false)
  const [, force] = useReducer((x: number) => x + 1, 0)

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
  // 最新 props(供 controller 构造时注册的 hook 读到当前回调)
  const propsRef = useRef(props)
  propsRef.current = props

  // ---- 实例化控制器(一次)。用 layout effect: 在 paint 前就绪,后续 rebuild 也是 layout,顺序确定。 ----
  useLayoutEffect(() => {
    const c = canvasRef.current
    const ra = renderAreaRef.current
    const sc = scrollerRef.current
    const sp = spacerRef.current
    if (!c || !ra || !sc || !sp || !ovMain.current || !ovFRow.current || !ovFCol.current || !ovCorner.current) return
    const controller = new ViewerController(
      {
        canvas: c,
        renderArea: ra,
        scroller: sc,
        spacer: sp,
        overlays: { main: ovMain.current, frow: ovFRow.current, fcol: ovFCol.current, corner: ovCorner.current },
      },
      {
        onRenderer: () => force(),
        onRenderTick: () => {}, // canvas 由 controller 直接画;React 不需每帧重渲
        onSelectionChange: () => force(),
        onCellClick: (row, col, text) => propsRef.current.onCellClick?.({ row, col, text }),
        onCellDblClick: (row, col, text) => propsRef.current.onCellDblClick?.({ row, col, text }),
        onHyperlink: (url, cell) => {
          propsRef.current.onHyperlinkClick?.({ url, cell })
          if (propsRef.current.openLinks !== false) window.open(url, '_blank', 'noopener')
        },
        onTooltip: (tip) => {
          tooltipRef.current = tip
          force()
        },
        onFindChange: () => force(),
        onFilterChange: () => force(),
      },
    )
    controller.fileName = propsRef.current.fileName
    controllerRef.current = controller

    const ro = new ResizeObserver(() => {
      controller.measure()
      controller.render()
    })
    ro.observe(ra)
    return () => {
      ro.disconnect()
      controller.dispose()
      controllerRef.current = null
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // ---- 载入 src ----
  useEffect(() => {
    if (props.src) load(props.src, props.transformModel)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.src])

  // ---- 文件名同步 ----
  useEffect(() => {
    if (controllerRef.current) controllerRef.current.fileName = props.fileName
  }, [props.fileName])

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

  // ---- 重建渲染器(工作簿 / 活动表 / 主题 / cellStyle 变化)。layout effect: 同步绘制,避免被晚到的 rebuild 清掉交互态 ----
  useLayoutEffect(() => {
    const controller = controllerRef.current
    if (!controller || !workbook) return
    const sheet: SheetModel | null = workbook.sheets[activeSheet] ?? workbook.sheets[0] ?? null
    if (!sheet) return
    controller.rebuild(sheet, workbook, zoom, { theme: props.theme, cellStyle: props.cellStyle })
    propsRef.current.onSheetChange?.({ index: activeSheet, name: sheet.name })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [workbook, activeSheet, props.theme, props.cellStyle])

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
      propsRef.current.onSelectionChange?.({ range: sel!, active })
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
      exportImage: (opts) => controllerRef.current!.exportImage(opts),
      downloadImage: (opts) => controllerRef.current!.downloadImage(opts),
      exportPdf: (opts) => controllerRef.current!.exportPdf(opts),
      downloadPdf: (opts) => controllerRef.current!.downloadPdf(opts),
      print: (opts) => controllerRef.current!.print(opts),
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
          <button onClick={() => void controllerRef.current?.downloadImage()}>导出 PNG</button>
          <button onClick={() => void controllerRef.current?.downloadPdf().catch((e) => propsRef.current.onError?.(String((e as Error)?.message ?? e)))}>
            导出 PDF
          </button>
          <select value={Math.round(zoom * 100)} onChange={(e) => setZoom(Number(e.target.value) / 100)} title="缩放">
            {[50, 75, 100, 125, 150, 200].map((p) => (
              <option key={p} value={p}>
                {p}%
              </option>
            ))}
          </select>
        </div>
      )}

      {workbook && (
        <div className="rxl-formula-bar">
          <span className="addr">{activeAddr || '—'}</span>
          <span className="fx">fx</span>
          <span className="content" title={formulaText}>
            {formulaText}
          </span>
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
        >
          <div className="rxl-spacer" ref={spacerRef} />
        </div>

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
