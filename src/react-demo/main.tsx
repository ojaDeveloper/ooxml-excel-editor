import { useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { ExcelViewer, type ExcelViewerHandle } from '@/react'
import { definePlugin } from '@/core/plugin'
import type { ExcelSource } from '@/core/loader'
import { demoSelectEditor } from '@/demo-shared/demo-editor'

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
  const ref = useRef<ExcelViewerHandle>(null)

  // 开发期把命令式句柄挂 window,供 e2e 取几何/读数据(与 Vue demo 的 __excelViewer 对齐)
  if (import.meta.env.DEV) {
    ;(window as unknown as { __excelViewerReact?: ExcelViewerHandle | null }).__excelViewerReact = ref.current
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh' }}>
      <div style={{ display: 'flex', gap: 8, padding: 8, borderBottom: '1px solid #e2e4e7', alignItems: 'center' }}>
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
            if (f) {
              setSrc(f)
              setFileName(f.name)
            }
          }}
        />
        <span style={{ color: '#888', fontSize: 13 }}>{fileName}</span>
        <label style={{ fontSize: 13 }} title="开启编辑模式(E0:闸门)">
          <input type="checkbox" checked={editMode} onChange={(e) => setEditMode(e.target.checked)} /> 编辑模式
        </label>
        {editMode && (
          <button
            onClick={() => {
              const sel = ref.current?.getSelection()
              if (sel) ref.current?.setStyle(sel, { font: { bold: true } })
            }}
            title="给选区加粗(E5:样式编辑)"
          >
            B 加粗选区
          </button>
        )}
        {editMode && (
          <button
            onClick={() => {
              const sel = ref.current?.getSelection()
              if (sel) ref.current?.insertRows(sel.top, 1)
            }}
            title="选区上方插入行(E7)"
          >
            ＋行
          </button>
        )}
        {editMode && (
          <button
            onClick={() => {
              const sel = ref.current?.getSelection()
              if (sel) ref.current?.deleteRows(sel.top, sel.bottom - sel.top + 1)
            }}
            title="删除选区行(E7)"
          >
            －行
          </button>
        )}
        <button onClick={() => ref.current?.downloadXlsx()} title="导出 .xlsx(E8:从模型重建)">
          ↓XLSX
        </button>
        <button onClick={() => ref.current?.downloadCsv()} title="导出 .csv(E8)">
          ↓CSV
        </button>
        <button onClick={() => ref.current?.downloadJson()} title="导出 .json(E8)">
          ↓JSON
        </button>
      </div>
      <div style={{ flex: 1, minHeight: 0 }}>
        <ExcelViewer
          ref={ref}
          src={src}
          fileName={fileName}
          plugins={[demoPlugin]}
          editable={editMode}
          recalc={editMode}
          readOnlyRanges={[{ top: 1, left: 0, bottom: 1, right: 4 }]}
          editor={demoSelectEditor}
          onCellChange={(p) => {
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
