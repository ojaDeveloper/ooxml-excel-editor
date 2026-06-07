/**
 * 内置导出进度遮罩(P1.5)—— React 等价物。居中模态 + stage 标签 + 进度条 + 取消。
 * 默认绑定;用方覆盖路径见 ExcelViewer.tsx 的 `exportProgress` prop 与 `renderExportProgress` render prop。
 */
import type { ExportProgress } from '@/core/progress'

const STAGE_LABEL: Record<string, string> = {
  render: '渲染中',
  compose: '合成中',
  paginate: '分页中',
  write: '写出文件',
  zip: 'zip 压缩',
  convert: '批量转换',
}

export function ExportProgressOverlay({
  state,
  busy,
  onCancel,
}: {
  state: ExportProgress | null
  busy: boolean
  onCancel: () => void
}) {
  if (!busy) return null
  const ratio = state?.ratio
  const label = state?.label || STAGE_LABEL[state?.stage ?? ''] || '处理中…'
  return (
    <div className="rxl-export-progress" role="dialog" aria-modal="true" aria-live="polite">
      <div className="card">
        <div className="title">{label}</div>
        <div className={'bar' + (ratio == null ? ' indeterminate' : '')}>
          {ratio != null ? <div className="fill" style={{ width: Math.round(ratio * 100) + '%' }} /> : null}
        </div>
        <div className="row">
          <span className="pct">{ratio != null ? Math.round(ratio * 100) + '%' : '正在处理…'}</span>
          <button className="cancel" onClick={onCancel} title="按 Esc 也可取消">取消</button>
        </div>
      </div>
    </div>
  )
}
