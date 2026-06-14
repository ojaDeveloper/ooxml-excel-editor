/**
 * 数据验证提示 UI(框架无关 DOM,三壳共用)。1.8.0 新增。
 *  - showError:非法输入提示。'stop' → 居中模态(标题 + 信息 + 知道了);'warning'/'information' → 顶部 toast 自动消失。
 *  - showInputBubble / hideInputBubble:选中带"输入提示"的格时,在格旁弹黄色气泡(WPS/Excel 同款)。
 * 纯 DOM + 内联样式,不依赖任何框架;跟 readonly-prompt-host 风格保持一致。
 */
export type ValidationErrorStyle = 'stop' | 'warning' | 'information'

const ICON: Record<ValidationErrorStyle, string> = { stop: '⛔', warning: '⚠️', information: 'ℹ️' }

export class ValidationPromptHost {
  private el: HTMLElement | null = null
  private toastEl: HTMLElement | null = null
  private toastTimer: ReturnType<typeof setTimeout> | null = null
  private bubbleEl: HTMLElement | null = null
  private cleanup: (() => void) | null = null
  private prevFocus: HTMLElement | null = null

  /** 非法输入提示。stop → 模态;warning/information → toast。 */
  showError(style: ValidationErrorStyle, message: string, title?: string): void {
    if (typeof document === 'undefined') return
    if (style === 'stop') this.showDialog(message, title)
    else this.showToast(style, message, title)
  }

  private showToast(style: ValidationErrorStyle, message: string, title?: string): void {
    this.closeToast()
    const t = document.createElement('div')
    t.className = 'ooxml-validation-toast'
    t.style.cssText = "position:fixed;top:16px;left:50%;transform:translateX(-50%);z-index:10001;background:#fff7e6;border:1px solid #ffd591;color:#874d00;border-radius:8px;padding:9px 16px;box-shadow:0 6px 24px rgba(0,0,0,.16);font:13px/1.5 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;max-width:80vw"
    t.textContent = `${ICON[style]} ${title ? title + ':' : ''}${message}`
    document.body.appendChild(t)
    this.toastEl = t
    this.toastTimer = setTimeout(() => this.closeToast(), 3600)
  }
  private closeToast(): void {
    if (this.toastTimer) { clearTimeout(this.toastTimer); this.toastTimer = null }
    this.toastEl?.remove()
    this.toastEl = null
  }

  private showDialog(message: string, title?: string): void {
    this.close()
    // 记住弹窗前的焦点(通常是单元格编辑器),关闭时还回去 → 用户接着改正不用重新点
    this.prevFocus = (document.activeElement as HTMLElement | null) ?? null
    const mask = document.createElement('div')
    mask.className = 'ooxml-validation-mask'
    mask.style.cssText = 'position:fixed;inset:0;z-index:10001;background:rgba(0,0,0,.32);display:flex;align-items:center;justify-content:center;'
    const card = document.createElement('div')
    card.style.cssText = "width:380px;max-width:92vw;background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.25);font:13px/1.5 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;color:#1f2329;overflow:hidden"
    card.innerHTML = [
      `<div style="display:flex;align-items:center;gap:8px;padding:13px 16px;border-bottom:1px solid #eef0f2;font-weight:600">⛔ ${escapeHtml(title || '输入不符合限制条件')}</div>`,
      `<div style="padding:16px;color:#475467">${escapeHtml(message)}</div>`,
      '<div style="display:flex;justify-content:flex-end;padding:12px 16px;background:#fafbfc;border-top:1px solid #eef0f2">',
      '<button data-ok style="height:30px;padding:0 16px;border:1px solid #1b7f4d;border-radius:5px;background:#21a366;color:#fff;cursor:pointer">知道了</button>',
      '</div>',
    ].join('')
    mask.appendChild(card)
    document.body.appendChild(mask)
    this.el = mask
    // Enter/Esc 关弹窗:阻断传播 + 默认,避免泄漏给后面的编辑器再次触发提交
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape' || e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); this.close() } }
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
    // 焦点还给弹窗前的元素(编辑器),让用户接着改
    const pf = this.prevFocus
    this.prevFocus = null
    if (pf && typeof pf.focus === 'function') setTimeout(() => pf.focus(), 0)
  }

  /** 选中带输入提示的格时,在视口坐标 (x,y) 处弹黄色气泡(x,y = 该格右下角附近)。 */
  showInputBubble(x: number, y: number, text: string, title?: string): void {
    if (typeof document === 'undefined') return
    this.hideInputBubble()
    const b = document.createElement('div')
    b.className = 'ooxml-validation-bubble'
    b.style.cssText = "position:fixed;z-index:9999;max-width:240px;background:#fffbe6;border:1px solid #ffe58f;color:#614700;border-radius:6px;padding:7px 10px;box-shadow:0 4px 16px rgba(0,0,0,.14);font:12px/1.5 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;pointer-events:none;white-space:pre-wrap;word-break:break-word"
    b.innerHTML = (title ? `<div style="font-weight:600;margin-bottom:2px">${escapeHtml(title)}</div>` : '') + escapeHtml(text)
    document.body.appendChild(b)
    // 先放上再按尺寸夹到视口内
    const vw = window.innerWidth, vh = window.innerHeight
    const rect = b.getBoundingClientRect()
    const left = Math.min(Math.max(4, x), vw - rect.width - 4)
    const top = Math.min(Math.max(4, y), vh - rect.height - 4)
    b.style.left = `${left}px`
    b.style.top = `${top}px`
    this.bubbleEl = b
  }
  hideInputBubble(): void {
    this.bubbleEl?.remove()
    this.bubbleEl = null
  }

  dispose(): void {
    this.close()
    this.closeToast()
    this.hideInputBubble()
  }
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string)
}
