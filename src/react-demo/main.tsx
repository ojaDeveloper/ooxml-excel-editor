import { useEffect, useLayoutEffect, useMemo, useReducer, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { ExcelViewer, type ExcelViewerHandle } from '@/react'
import { definePlugin } from '@/core/plugin'
import type { ExcelSource } from '@/core/loader'
import { demoSelectEditor } from '@/demo-shared/demo-editor'
import './demo-bar.css'

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
      <button key={it.id} className="rxl-demo-btn" title={it.title} onClick={() => { it.onClick(); onAfter?.() }}>
        {it.label}
      </button>
    )
  }
  if (it.kind === 'color') {
    return (
      <label key={it.id} className="rxl-demo-lb" title={it.title}>
        {it.label}
        <input type="color" value={it.getColor()} onChange={(e) => it.onColor(e.target.value)} />
      </label>
    )
  }
  return (
    <label key={it.id} className="rxl-demo-lb" title={it.title}>
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
  const [fileName, setFileName] = useState('')
  const [editMode, setEditMode] = useState(false) // E0: 编辑模式闸门
  const [fit, setFit] = useState<'fill' | 'contain' | 'cover'>('contain') // WPS 内嵌图贴合方式(默认 contain 同 WPS)
  const [, bumpSel] = useReducer((x: number) => x + 1, 0) // 选区/内容变 → 重渲(颜色回显)
  // 稳定引用:demo 因 bumpSel 频繁重渲,这些 prop 若每次新建数组会让壳的 effect 反复重跑(清掉选区)
  const readOnlyRanges = useMemo(() => [{ top: 1, left: 0, bottom: 1, right: 4 }], [])
  const plugins = useMemo(() => [demoPlugin], [])
  const ref = useRef<ExcelViewerHandle>(null)

  // 开发期把命令式句柄挂 window,供 e2e 取几何/读数据(与 Vue demo 的 __excelViewer 对齐)
  if (import.meta.env.DEV) {
    ;(window as unknown as { __excelViewerReact?: ExcelViewerHandle | null }).__excelViewerReact = ref.current
  }

  // demo 演示按钮(放不下自动收进「更多」)
  const items: DemoItem[] = []
  if (src) {
    if (editMode) {
      items.push(
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
      <div ref={barRef} className="rxl-demo-bar">
        <div ref={fixedRef} className="rxl-demo-fixed">
          <strong style={{ fontSize: 14 }}>React 版 ExcelViewer(共用 core)</strong>
          <button
            onClick={() => {
              setSrc(import.meta.env.BASE_URL + 'sample.xlsx')
              setFileName('sample.xlsx')
            }}
          >
            加载示例
          </button>
          <input
            type="file"
            accept=".xlsx,.xlsm"
            onChange={(e) => {
              const f = e.target.files?.[0]
              if (f) { setSrc(f); setFileName(f.name) }
            }}
          />
          <span style={{ color: '#888', fontSize: 13 }}>{fileName}</span>
          <label style={{ fontSize: 13 }} title="开启编辑模式(E0:闸门)">
            <input type="checkbox" checked={editMode} onChange={(e) => setEditMode(e.target.checked)} /> 编辑模式
          </label>
        </div>
        <div className="rxl-demo-grow" />
        {/* 隐藏测量行 */}
        <div ref={measureRef} className="rxl-demo-measure" aria-hidden="true">
          {items.map((it) => renderItem(it))}
        </div>
        {visibleItems.map((it) => renderItem(it))}
        {overflowItems.length ? (
          <div className="rxl-demo-more">
            <button className={'rxl-demo-btn' + (moreOpen ? ' open' : '')} title="更多" onClick={() => setMoreOpen(!moreOpen)}>
              ⋯ 更多
            </button>
            {moreOpen ? (
              <div className="rxl-demo-pop">
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
          fileName={fileName}
          plugins={plugins}
          editable={editMode}
          cellImageFit={fit}
          recalc={editMode}
          readOnlyRanges={readOnlyRanges}
          editor={demoSelectEditor}
          toolbar={['find', 'filter', 'clear-filter', 'separator', 'copy', 'wrap-text', 'image-tools', 'freeze', 'separator', 'template', 'separator', 'zoom', 'export']}
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
    </div>
  )
}

createRoot(document.getElementById('root')!).render(<Demo />)
