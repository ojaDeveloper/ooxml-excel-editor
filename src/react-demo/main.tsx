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
      </div>
      <div style={{ flex: 1, minHeight: 0 }}>
        <ExcelViewer
          ref={ref}
          src={src}
          fileName={fileName}
          plugins={[demoPlugin]}
          editable={editMode}
          readOnlyRanges={[{ top: 1, left: 0, bottom: 1, right: 4 }]}
          editor={demoSelectEditor}
          onCellChange={(p) => {
            if (import.meta.env.DEV) (window as unknown as { __lastCellChange?: unknown }).__lastCellChange = p
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
