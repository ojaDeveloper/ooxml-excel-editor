/**
 * 只读提醒 UI(框架无关 DOM,三壳共用)。粘贴撞只读格时,按配置:
 *  - 'dialog'(默认):弹窗列出**具体哪些格**只读(A1 引用),让用户明确知道哪部分没粘上;
 *  - 'toast':顶部气泡,简短提示 + 自动消失;
 *  - 'none':不弹内置 UI(只走 permission-denied 事件,消费方自处理)。
 * 逐格精确:cells 是被跳过的只读格列表。
 */
export type ReadOnlyPromptMode = 'dialog' | 'toast' | 'none'

function colLabel(col: number): string {
  let s = ''
  let n = col
  do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1 } while (n >= 0)
  return s
}
const a1 = (c: { row: number; col: number }) => `${colLabel(c.col)}${c.row + 1}`

export class ReadOnlyPromptHost {
  private el: HTMLElement | null = null
  private toastEl: HTMLElement | null = null
  private toastTimer: ReturnType<typeof setTimeout> | null = null
  private cleanup: (() => void) | null = null

  show(mode: ReadOnlyPromptMode, message: string, cells: Array<{ row: number; col: number }>): void {
    if (typeof document === 'undefined' || mode === 'none') return
    if (mode === 'toast') this.showToast(message)
    else this.showDialog(message, cells)
  }

  private showToast(message: string): void {
    this.closeToast()
    const t = document.createElement('div')
    t.className = 'ooxml-readonly-toast'
    t.style.cssText = "position:fixed;top:16px;left:50%;transform:translateX(-50%);z-index:10001;background:#fff7e6;border:1px solid #ffd591;color:#874d00;border-radius:8px;padding:9px 16px;box-shadow:0 6px 24px rgba(0,0,0,.16);font:13px/1.5 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;max-width:80vw"
    t.textContent = '🔒 ' + message
    document.body.appendChild(t)
    this.toastEl = t
    this.toastTimer = setTimeout(() => this.closeToast(), 3600)
  }
  private closeToast(): void {
    if (this.toastTimer) { clearTimeout(this.toastTimer); this.toastTimer = null }
    this.toastEl?.remove()
    this.toastEl = null
  }

  private showDialog(message: string, cells: Array<{ row: number; col: number }>): void {
    this.close()
    const refs = cells.map(a1)
    const shown = refs.slice(0, 60)
    const more = refs.length - shown.length
    const mask = document.createElement('div')
    mask.className = 'ooxml-readonly-mask'
    mask.style.cssText = 'position:fixed;inset:0;z-index:10001;background:rgba(0,0,0,.32);display:flex;align-items:center;justify-content:center;'
    const card = document.createElement('div')
    card.style.cssText = "width:420px;max-width:92vw;background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.25);font:13px/1.5 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;color:#1f2329;overflow:hidden"
    card.innerHTML = [
      '<div style="display:flex;align-items:center;gap:8px;padding:13px 16px;border-bottom:1px solid #eef0f2;font-weight:600">🔒 只读,操作已跳过</div>',
      '<div style="padding:14px 16px 6px">',
      `<div style="margin-bottom:10px;color:#475467">${escapeHtml(message)}</div>`,
      refs.length
        ? `<div style="color:#667085;margin-bottom:4px">只读格(${refs.length} 个):</div><div style="max-height:180px;overflow:auto;background:#f7f8fa;border:1px solid #eef0f2;border-radius:6px;padding:8px;line-height:1.9">${shown.map((r) => `<span style="display:inline-block;background:#fff;border:1px solid #e4e7ec;border-radius:4px;padding:0 6px;margin:0 4px 4px 0">${r}</span>`).join('')}${more > 0 ? `<span style="color:#98a2b3">…另 ${more} 个</span>` : ''}</div>`
        : '',
      '</div>',
      '<div style="display:flex;justify-content:flex-end;padding:12px 16px;background:#fafbfc;border-top:1px solid #eef0f2">',
      '<button data-ok style="height:30px;padding:0 16px;border:1px solid #1b7f4d;border-radius:5px;background:#21a366;color:#fff;cursor:pointer">知道了</button>',
      '</div>',
    ].join('')
    mask.appendChild(card)
    document.body.appendChild(mask)
    this.el = mask
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape' || e.key === 'Enter') this.close() }
    card.querySelector('[data-ok]')?.addEventListener('click', () => this.close())
    mask.addEventListener('mousedown', (e) => { if (e.target === mask) this.close() })
    document.addEventListener('keydown', onKey, true)
    this.cleanup = () => document.removeEventListener('keydown', onKey, true)
  }

  close(): void {
    this.cleanup?.()
    this.cleanup = null
    this.el?.remove()
    this.el = null
  }
  dispose(): void {
    this.close()
    this.closeToast()
  }
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string)
}
