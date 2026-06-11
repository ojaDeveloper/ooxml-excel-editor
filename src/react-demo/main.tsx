import { useEffect, useLayoutEffect, useMemo, useReducer, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { ExcelViewer, type ExcelViewerHandle } from '@/react'
import { definePlugin } from '@/core/plugin'
import type { ExcelSource } from '@/core/loader'
import { demoSelectEditor } from '@/demo-shared/demo-editor'

import '../demo-shared/demo-bar.css'

// ---------------- demo 顶栏溢出工具(只用于本 demo;不污染组件) ----------------
type DemoItem =
  | { id: string; kind: 'btn'; label: string; title?: string; onClick: () => void }
  | { id: string; kind: 'color'; label: string; title?: string; getColor: () => string; onColor: (c: string) => void }
  | {
      id: string
      kind: 'select'
      label: string
      title?: string
      value: () => string
      options: { value: string; label: string }[]
      onChange: (v: string) => void
    }

function renderItem(it: DemoItem, onAfter?: () => void): JSX.Element {
  if (it.kind === 'btn') {
    return (
      <button key={it.id} className="sample-btn" title={it.title} onClick={() => { it.onClick(); onAfter?.() }}>
        {it.label}
      </button>
    )
  }
  if (it.kind === 'color') {
    return (
      <label key={it.id} className="sample-label" title={it.title}>
        {it.label}
        <input type="color" value={it.getColor()} onChange={(e) => it.onColor(e.target.value)} />
      </label>
    )
  }
  return (
    <label key={it.id} className="sample-label" title={it.title}>
      {it.label}
      <select value={it.value()} onChange={(e) => { it.onChange(e.target.value); onAfter?.() }}>
        {it.options.map((o) => (
          <option key={o.value} value={o.value}>
            {o.label}
          </option>
        ))}
      </select>
    </label>
  )
}

// 跨框架插件:同一份 definePlugin 在 Vue / React 都能用。overlay 返回 DOM(框架无关)。
const demoPlugin = definePlugin({
  name: 'react-demo-plugin',
  cellStyle: (c) => (typeof c.raw === 'number' && c.raw < 0 ? { font: { color: '#d00' } } : undefined),
  toolbar: [{ id: 'plugin-hello', label: '🔌 插件按钮', onClick: (v) => v.setSelection({ top: 1, left: 0, bottom: 1, right: 0 }) }],
  overlay: ({ rectOf }) => {
    const r = rectOf(2, 1) // B3
    if (!r) return null
    const el = document.createElement('div')
    el.className = 'plugin-badge'
    el.textContent = '🎯'
    Object.assign(el.style, {
      position: 'absolute',
      left: r.x + r.w - 14 + 'px',
      top: r.y - 2 + 'px',
      fontSize: '12px',
      pointerEvents: 'none',
    })
    return el
  },
})

function Demo() {
  const [src, setSrc] = useState<ExcelSource | undefined>(undefined)
  const [jsonItems, setJsonItems] = useState<Array<Record<string, unknown>> | null>(null)
  const [fileName, setFileName] = useState('')
  const [editMode, setEditMode] = useState(false) // E0: 编辑模式闸门
  const [fit, setFit] = useState<'fill' | 'contain' | 'cover'>('contain') // WPS 内嵌图贴合方式(默认 contain 同 WPS)
  const [highlightReadOnly, setHighlightReadOnly] = useState(false)
  const [editableTargetsApplied, setEditableTargetsApplied] = useState<Array<{ row: number } | { col: number } | { row: number; col: number }> | undefined>(undefined)
  const [editTargetsDialogOpen, setEditTargetsDialogOpen] = useState(false)
  const [editTargetsCells, setEditTargetsCells] = useState<Set<string>>(new Set())
  const [editTargetsRows, setEditTargetsRows] = useState<Set<number>>(new Set())
  const [editTargetsCols, setEditTargetsCols] = useState<Set<number>>(new Set())
  const [, bumpSel] = useReducer((x: number) => x + 1, 0) // 选区/内容变 → 重渲(颜色回显)

  function openEditTargetsDialog() {
    const cells = new Set<string>()
    const rows = new Set<number>()
    const cols = new Set<number>()
    for (const t of editableTargetsApplied ?? []) {
      const tr = (t as { row?: number }).row
      const tc = (t as { col?: number }).col
      if (typeof tr === 'number' && typeof tc === 'number') cells.add(`${tr}:${tc}`)
      else if (typeof tr === 'number') rows.add(tr)
      else if (typeof tc === 'number') cols.add(tc)
    }
    setEditTargetsCells(cells); setEditTargetsRows(rows); setEditTargetsCols(cols)
    setEditTargetsDialogOpen(true)
  }
  function toggleEditTargetCell(r: number, c: number) {
    const k = `${r}:${c}`
    setEditTargetsCells((prev) => { const next = new Set(prev); if (next.has(k)) next.delete(k); else next.add(k); return next })
  }
  function toggleEditTargetRow(r: number) {
    setEditTargetsRows((prev) => { const next = new Set(prev); if (next.has(r)) next.delete(r); else next.add(r); return next })
  }
  function toggleEditTargetCol(c: number) {
    setEditTargetsCols((prev) => { const next = new Set(prev); if (next.has(c)) next.delete(c); else next.add(c); return next })
  }
  const isCellInDraft = (r: number, c: number) => editTargetsCells.has(`${r}:${c}`) || editTargetsRows.has(r) || editTargetsCols.has(c)
  function applyEditTargets() {
    const arr: Array<{ row: number } | { col: number } | { row: number; col: number }> = []
    for (const r of editTargetsRows) arr.push({ row: r })
    for (const c of editTargetsCols) arr.push({ col: c })
    for (const k of editTargetsCells) {
      const [r, c] = k.split(':').map(Number)
      if (editTargetsRows.has(r) || editTargetsCols.has(c)) continue
      arr.push({ row: r, col: c })
    }
    setEditableTargetsApplied(arr)
    setEditTargetsDialogOpen(false)
  }
  function clearEditTargets() {
    setEditableTargetsApplied(undefined)
    setEditTargetsDialogOpen(false)
  }
  const colLetter = (c: number) => {
    let s = ''; let n = c
    while (true) { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; if (n < 0) break }
    return s
  }
  const previewCellText = (r: number, c: number) => {
    const v = ref.current?.getCellText(r, c) ?? ''
    return v.length > 6 ? v.slice(0, 6) + '…' : v
  }
  // 稳定引用:demo 因 bumpSel 频繁重渲,这些 prop 若每次新建数组会让壳的 effect 反复重跑(清掉选区)
  const readOnlyRanges = useMemo(() => [{ top: 1, left: 0, bottom: 1, right: 4 }], [])
  const plugins = useMemo(() => [demoPlugin], [])
  const ref = useRef<ExcelViewerHandle>(null)

  // 跟 Vue 3 demo 同款的"演示功能" — 加载 JSON / PDF 水印 / 数据→JSON
  function loadJsonSample() {
    setSrc(undefined)
    setFileName('订单数据')
    setJsonItems([
      { name: '笔记本电脑', price: 5999, qty: 1, amount: 5999, note: '商务款' },
      { name: '机械键盘', price: 399, qty: 2, amount: 798, note: '青轴' },
      { name: '显示器', price: 1299, qty: 2, amount: 2598, note: '27寸 2K' },
      { name: '鼠标', price: 89, qty: 5, amount: 445, note: '无线' },
      { name: '耳机', price: 599, qty: 3, amount: 1797, note: '降噪' },
    ])
  }
  async function exportPdfWithWatermark() {
    try {
      await ref.current?.downloadPdf({
        target: 'all',
        beforeRenderPage: (ctx: any) => {
          const { doc, pageIndex, pageCount, pageWidth, pageHeight, margin, sheetName } = ctx
          doc.setFontSize(9); doc.setTextColor(120)
          doc.text(`${sheetName}`, margin.left, pageHeight - 5)
          doc.text(`第 ${pageIndex + 1} / ${pageCount} 页`, pageWidth - margin.right, pageHeight - 5, { align: 'right' })
          doc.setFontSize(56); doc.setTextColor(230)
          doc.text('PREVIEW', pageWidth / 2, pageHeight / 2, { align: 'center', angle: 30 })
        },
      })
    } catch (e) { console.error('PDF 水印导出失败:', e) }
  }
  function showSheetJSON() {
    const json = ref.current?.getSheetJSON({ headerRow: 1 }) ?? []
    navigator.clipboard?.writeText(JSON.stringify(json, null, 2)).catch(() => {})
    alert(`${json.length} 行已复制为 JSON · 首行: ${JSON.stringify(json[0] ?? {})}`.slice(0, 200))
  }
  function jumpToLastRow() {
    const viewer = ref.current
    const wb = viewer?.getWorkbook()
    if (!viewer || !wb) return
    const sheet = wb.sheets[viewer.getActiveSheet()]
    const row = Math.max(0, sheet.dimension.rows - 1)
    viewer.scrollToCell(row, 0, { select: true })
    alert(`已跳到末行 A${row + 1}`)
  }

  // 开发期把命令式句柄挂 window,供 e2e 取几何/读数据(与 Vue demo 的 __excelViewer 对齐)
  if (import.meta.env.DEV) {
    ;(window as unknown as { __excelViewerReact?: ExcelViewerHandle | null }).__excelViewerReact = ref.current
  }

  // demo 演示按钮(放不下自动收进「更多」)
  const items: DemoItem[] = []
  if (src || jsonItems) {
    items.push(
      { id: 'pdf-watermark', kind: 'btn', label: 'PDF(页码+水印)', title: '演示 beforeRenderPage 钩子', onClick: () => void exportPdfWithWatermark() },
      { id: 'sheet-json', kind: 'btn', label: '数据→JSON', title: '演示数据读取 API getSheetJSON', onClick: showSheetJSON },
      { id: 'jump-last-row', kind: 'btn', label: '跳到末行', title: '演示 scrollToCell(row,col,{select:true}) 导航 API', onClick: jumpToLastRow },
    )
  }
  if (src) {
    if (editMode) {
      items.push(
        { id: 'edit-targets', kind: 'btn', label: editableTargetsApplied ? `可编辑 (${editableTargetsApplied.length})` : '设置可编辑', title: '白名单模式: 点选要可编辑的格 / 行 / 列, 应用后只这些可编辑', onClick: openEditTargetsDialog },
        { id: 'highlight-readonly', kind: 'btn', label: highlightReadOnly ? '✓ 高亮只读' : '高亮只读', title: '把只读格套浅灰底', onClick: () => setHighlightReadOnly(!highlightReadOnly) },
        { id: 'bold', kind: 'btn', label: 'B 加粗选区', title: '给选区加粗(E5)', onClick: () => { const s = ref.current?.getSelection(); if (s) ref.current?.setStyle(s, { font: { bold: true } }) } },
        { id: 'merge', kind: 'btn', label: '合并', title: '合并选区(G1)', onClick: () => { const s = ref.current?.getSelection(); if (s) ref.current?.mergeCells(s) } },
        { id: 'unmerge', kind: 'btn', label: '拆分', title: '拆分选区(G1)', onClick: () => { const s = ref.current?.getSelection(); if (s) ref.current?.unmergeCells(s) } },
        { id: 'fill', kind: 'color', label: '背景', title: '背景填充色', getColor: () => ref.current?.getActiveFillColor() ?? '#FFFFFF', onColor: (c) => { ref.current?.setSelectionFill(c); bumpSel() } },
        { id: 'font', kind: 'color', label: '字体', title: '字体颜色', getColor: () => ref.current?.getActiveFontColor() ?? '#000000', onColor: (c) => { ref.current?.setSelectionFontColor(c); bumpSel() } },
        { id: 'clear-fill', kind: 'btn', label: '清除填充', onClick: () => { ref.current?.setSelectionFill(null); bumpSel() } },
        { id: 'wrap', kind: 'btn', label: '自动换行', title: '自动换行(WPS 风格 toggle)', onClick: () => { ref.current?.toggleWrapTextOnSelection(); bumpSel() } },
        { id: 'embed-all', kind: 'btn', label: '整表嵌入', title: 'WPS 浮动→嵌入(DISPIMG)', onClick: () => { const n = ref.current?.convertAllImagesToCells() ?? 0; if (!n) alert('没有可嵌入的浮动图') } },
        { id: 'cell-to-float', kind: 'btn', label: '格→图', title: '内嵌图→浮动图', onClick: () => { const s = ref.current?.getSelection(); if (s) ref.current?.convertCellImageToFloat(s.top, s.left) } },
        { id: 'ins-row', kind: 'btn', label: '＋行', title: '选区上方插入行', onClick: () => { const s = ref.current?.getSelection(); if (s) ref.current?.insertRows(s.top, 1) } },
        { id: 'del-row', kind: 'btn', label: '－行', title: '删除选区行', onClick: () => { const s = ref.current?.getSelection(); if (s) ref.current?.deleteRows(s.top, s.bottom - s.top + 1) } },
      )
    }
    items.push(
      { id: 'fit', kind: 'select', label: '贴合', title: 'WPS 内嵌图贴合方式', value: () => fit, options: [
        { value: 'contain', label: 'contain 等比(同 WPS)' }, { value: 'fill', label: 'fill 铺满' }, { value: 'cover', label: 'cover 裁剪' },
      ], onChange: (v) => setFit(v as 'fill' | 'contain' | 'cover') },
      { id: 'dl-xlsx', kind: 'btn', label: '↓XLSX', title: '导出 .xlsx', onClick: () => ref.current?.downloadXlsx() },
      { id: 'dl-csv', kind: 'btn', label: '↓CSV', title: '导出 .csv', onClick: () => ref.current?.downloadCsv() },
      { id: 'dl-json', kind: 'btn', label: '↓JSON', title: '导出 .json', onClick: () => ref.current?.downloadJson() },
    )
  }

  // 测量 + 溢出
  const barRef = useRef<HTMLDivElement | null>(null)
  const fixedRef = useRef<HTMLDivElement | null>(null)
  const measureRef = useRef<HTMLDivElement | null>(null)
  const [widths, setWidths] = useState<number[]>([])
  const [barContentW, setBarContentW] = useState(0)
  const [moreOpen, setMoreOpen] = useState(false)
  // 稳定 deps:items 数组每帧 new,直接放进 deps 会触发 setState → 新 items → 再触发 → 无限循环。
  // 用项数+id 串拼成字符串作为变更指纹。
  const itemsKey = items.map((i) => i.id).join('|')
  useLayoutEffect(() => {
    const m = measureRef.current
    if (!m) return
    const nextWidths = Array.from(m.children).map((c) => (c as HTMLElement).offsetWidth)
    setWidths((prev) => (prev.length === nextWidths.length && prev.every((v, i) => v === nextWidths[i]) ? prev : nextWidths))
    const bar = barRef.current
    const fixed = fixedRef.current
    const fixedW = fixed ? fixed.getBoundingClientRect().width : 0
    const nextW = Math.max(0, (bar?.clientWidth ?? 0) - fixedW - 24)
    setBarContentW((prev) => (prev === nextW ? prev : nextW))
  }, [itemsKey, fileName])
  useEffect(() => {
    const ro = new ResizeObserver(() => {
      const bar = barRef.current
      const fixed = fixedRef.current
      const fixedW = fixed ? fixed.getBoundingClientRect().width : 0
      setBarContentW(Math.max(0, (bar?.clientWidth ?? 0) - fixedW - 24))
    })
    if (barRef.current) ro.observe(barRef.current)
    const onDoc = (e: MouseEvent) => {
      if (barRef.current && !barRef.current.contains(e.target as Node)) setMoreOpen(false)
    }
    document.addEventListener('mousedown', onDoc)
    return () => { ro.disconnect(); document.removeEventListener('mousedown', onDoc) }
  }, [])
  const MORE_W = 64
  const GAP = 6
  const visibleCount = (() => {
    if (!barContentW || widths.length !== items.length) return items.length
    let sum = 0, fitsAll = true
    for (let i = 0; i < items.length; i++) { sum += widths[i] + GAP; if (sum > barContentW) { fitsAll = false; break } }
    if (fitsAll) return items.length
    let s = MORE_W, n = 0
    for (let i = 0; i < items.length; i++) { s += widths[i] + GAP; if (s > barContentW) break; n++ }
    return Math.max(0, n)
  })()
  const visibleItems = items.slice(0, visibleCount)
  const overflowItems = items.slice(visibleCount)

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh' }}>
      {/* demo 顶栏:固定区(标题/选 xlsx/加载示例/编辑模式)+ 演示按钮区(自动溢出收进「⋯ 更多」) */}
      <div ref={barRef} className="app-bar">
        <div ref={fixedRef} className="app-bar-fixed">
          <strong>OOXML Excel 预览器</strong>
          <span className="sub">React · Canvas 高保真</span>
          <label className="file-btn">
            选择 .xlsx
            <input
              type="file"
              accept=".xlsx,.xlsm"
              hidden
              onChange={(e) => {
                const f = e.target.files?.[0]
                if (f) { setJsonItems(null); setSrc(f); setFileName(f.name) }
              }}
            />
          </label>
          <button
            className="sample-btn"
            onClick={() => {
              setJsonItems(null)
              setSrc(import.meta.env.BASE_URL + 'sample.xlsx')
              setFileName('sample.xlsx')
            }}
          >
            加载示例
          </button>
          <button className="sample-btn" onClick={loadJsonSample} title="加载一个 JSON 数据源演示;然后用工具栏「模板」导入 .xlsx 看模板效果">
            JSON 示例
          </button>
          {(src || jsonItems) && (
            <label className="edit-toggle" title="开启编辑模式(E0:闸门)">
              <input type="checkbox" checked={editMode} onChange={(e) => setEditMode(e.target.checked)} /> 编辑模式
            </label>
          )}
        </div>
        <div className="grow" />
        {/* 隐藏测量行 */}
        <div ref={measureRef} className="app-bar-measure" aria-hidden="true">
          {items.map((it) => renderItem(it))}
        </div>
        {visibleItems.map((it) => renderItem(it))}
        {overflowItems.length ? (
          <div className="more-wrap">
            <button className={'sample-btn more-btn' + (moreOpen ? ' open' : '')} title="更多" onClick={() => setMoreOpen(!moreOpen)}>
              ⋯ 更多
            </button>
            {moreOpen ? (
              <div className="more-pop">
                {overflowItems.map((it) => renderItem(it, () => setMoreOpen(false)))}
              </div>
            ) : null}
          </div>
        ) : null}
      </div>
      <div style={{ flex: 1, minHeight: 0 }}>
        <ExcelViewer
          ref={ref}
          src={src}
          workbook={jsonItems ?? undefined}
          fileName={fileName}
          plugins={plugins}
          editable={editMode}
          pivotTable={true}
          cellImageFit={fit}
          recalc={editMode}
          readOnlyRanges={readOnlyRanges}
          editableTargets={editableTargetsApplied}
          readOnlyCellStyle={highlightReadOnly}
          editor={demoSelectEditor}
          toolbar={['find', 'filter', 'sort', 'clear-filter', 'separator', 'copy', 'pivot-table', 'wrap-text', 'image-tools', 'freeze', 'separator', 'template', 'separator', 'zoom', 'export']}
          onSelectionChange={() => bumpSel()}
          onCellChange={(p) => {
            bumpSel() // 颜色回显随内容/样式变更刷新
            if (import.meta.env.DEV) (window as unknown as { __lastCellChange?: unknown }).__lastCellChange = p
          }}
          onDimChange={(p) => {
            if (import.meta.env.DEV) (window as unknown as { __lastDimChange?: unknown }).__lastDimChange = p
          }}
          onDirtyChange={(p) => {
            if (import.meta.env.DEV) (window as unknown as { __lastDirtyChange?: unknown }).__lastDirtyChange = p
          }}
          onImageChange={(p) => {
            if (import.meta.env.DEV) (window as unknown as { __lastImageChange?: unknown }).__lastImageChange = p
          }}
          onStructChange={(p) => {
            if (import.meta.env.DEV) (window as unknown as { __lastStructChange?: unknown }).__lastStructChange = p
          }}
          onRendered={() => {
            // ref.current 此时已就绪,再挂一次保证 e2e 拿到
            if (import.meta.env.DEV) (window as unknown as { __excelViewerReact?: ExcelViewerHandle | null }).__excelViewerReact = ref.current
          }}
        />
      </div>
      {editTargetsDialogOpen && (
        <div className="edit-targets-overlay" onClick={(e) => { if (e.target === e.currentTarget) setEditTargetsDialogOpen(false) }}>
          <div className="edit-targets-dialog">
            <header>
              <h3>设置可编辑单元格 (白名单)</h3>
              <p className="hint">
                点击单元格 = 该格可编辑;点击列标题 (A/B/C…) = 整列可编辑;点击行号 = 整行可编辑.
                应用后,只有勾选的位置可编辑,其它全部只读. 关闭白名单 = 恢复默认 (整表可编辑).
              </p>
            </header>
            <div className="edit-targets-grid">
              <table>
                <thead>
                  <tr>
                    <th className="corner">#</th>
                    {Array.from({ length: 8 }).map((_, c) => (
                      <th key={c}
                          className={editTargetsCols.has(c) ? 'picked' : ''}
                          onClick={() => toggleEditTargetCol(c)}
                          title={`整列 ${colLetter(c)} 可编辑`}>{colLetter(c)}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {Array.from({ length: 12 }).map((_, r) => (
                    <tr key={r}>
                      <th className={editTargetsRows.has(r) ? 'picked' : ''}
                          onClick={() => toggleEditTargetRow(r)}
                          title={`整行 ${r + 1} 可编辑`}>{r + 1}</th>
                      {Array.from({ length: 8 }).map((_, c) => {
                        const cls: string[] = []
                        if (isCellInDraft(r, c)) cls.push('picked')
                        if (editTargetsRows.has(r) || editTargetsCols.has(c)) cls.push('row-col-hit')
                        return (
                          <td key={c} className={cls.join(' ')}
                              onClick={() => toggleEditTargetCell(r, c)}
                              title={`R${r + 1}C${c + 1}`}>{previewCellText(r, c) || '·'}</td>
                        )
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <footer>
              <span className="count-hint">
                已选: {editTargetsCells.size} 单格 / {editTargetsRows.size} 整行 / {editTargetsCols.size} 整列
              </span>
              <button className="dlg-btn ghost" onClick={() => setEditTargetsDialogOpen(false)}>取消</button>
              <button className="dlg-btn ghost" onClick={clearEditTargets} title="移除白名单, 恢复默认 (全可编辑)">关闭白名单</button>
              <button className="dlg-btn primary" onClick={applyEditTargets}>应用</button>
            </footer>
          </div>
        </div>
      )}
    </div>
  )
}

createRoot(document.getElementById('root')!).render(<Demo />)
