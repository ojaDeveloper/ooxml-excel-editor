import { useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { ExcelViewer, type ExcelViewerHandle } from '@/react'
import type { ExcelSource } from '@/core/loader'

function Demo() {
  const [src, setSrc] = useState<ExcelSource | undefined>(undefined)
  const [fileName, setFileName] = useState('')
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
      </div>
      <div style={{ flex: 1, minHeight: 0 }}>
        <ExcelViewer
          ref={ref}
          src={src}
          fileName={fileName}
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
