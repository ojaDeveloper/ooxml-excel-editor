/**
 * 批注编辑对话框(框架无关 DOM,三壳共用一份)。1.11.0 新增。
 * 右键「插入/编辑批注」打开:多行文本框 + 确定/删除/取消。确定回调批注文本(空 = 删除)。
 */
export interface CommentDialogOptions {
  cellRef: string // A1 引用,标题显示
  current: string
  onApply: (text: string) => void
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string)
}
const BTN = 'height:30px;padding:0 14px;border-radius:5px;cursor:pointer;font:inherit'

export class CommentDialogHost {
  private el: HTMLElement | null = null
  private cleanup: (() => void) | null = null

  show(opts: CommentDialogOptions): void {
    if (typeof document === 'undefined') return
    this.close()
    const mask = document.createElement('div')
    mask.className = 'ooxml-comment-mask'
    mask.style.cssText = 'position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.32);display:flex;align-items:center;justify-content:center;'
    const card = document.createElement('div')
    card.className = 'ooxml-comment-card'
    card.style.cssText = "width:380px;max-width:92vw;background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.25);font:13px/1.5 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;color:#1f2329;overflow:hidden"
    card.innerHTML = [
      `<div style="display:flex;align-items:center;justify-content:space-between;padding:13px 16px;border-bottom:1px solid #eef0f2;font-weight:600">${opts.current ? '编辑' : '插入'}批注 · ${escapeHtml(opts.cellRef)}<button data-close style="border:0;background:none;font-size:20px;line-height:1;color:#8a8f98;cursor:pointer">×</button></div>`,
      '<div style="padding:14px 16px">',
      `<textarea data-text rows="5" placeholder="输入批注…" style="width:100%;box-sizing:border-box;border:1px solid #d0d5dd;border-radius:6px;padding:8px;font:inherit;resize:vertical">${escapeHtml(opts.current)}</textarea>`,
      '</div>',
      '<div style="display:flex;justify-content:space-between;gap:8px;padding:12px 16px;background:#fafbfc;border-top:1px solid #eef0f2">',
      `<button data-del style="${BTN};border:1px solid #fda29b;background:#fff;color:#d92d20"${opts.current ? '' : ' disabled style="display:none"'}>删除批注</button>`,
      `<div style="display:flex;gap:8px;margin-left:auto"><button data-cancel style="${BTN};border:1px solid #d0d5dd;background:#fff;color:#1f2329">取消</button><button data-ok style="${BTN};border:1px solid #1b7f4d;background:#21a366;color:#fff">确定</button></div>`,
      '</div>',
    ].join('')
    mask.appendChild(card)
    document.body.appendChild(mask)
    this.el = mask
    const ta = card.querySelector('[data-text]') as HTMLTextAreaElement
    const apply = (text: string) => { opts.onApply(text); this.close() }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') { e.preventDefault(); this.close() }
      else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); apply(ta.value) }
    }
    card.querySelector('[data-close]')?.addEventListener('click', () => this.close())
    card.querySelector('[data-cancel]')?.addEventListener('click', () => this.close())
    card.querySelector('[data-ok]')?.addEventListener('click', () => apply(ta.value))
    card.querySelector('[data-del]')?.addEventListener('click', () => apply(''))
    mask.addEventListener('mousedown', (e) => { if (e.target === mask) this.close() })
    document.addEventListener('keydown', onKey, true)
    this.cleanup = () => document.removeEventListener('keydown', onKey, true)
    setTimeout(() => { ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length) }, 0)
  }

  close(): void {
    this.cleanup?.()
    this.cleanup = null
    this.el?.remove()
    this.el = null
  }
  dispose(): void { this.close() }
}
