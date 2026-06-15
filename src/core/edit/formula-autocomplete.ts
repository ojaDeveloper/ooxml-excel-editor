/**
 * 公式自动补全(框架无关 DOM)。1.14.0 新增。挂到单元格编辑器 / 公式栏的 <textarea>:
 * 输入 `=SU` 时在下方弹函数名列表(带签名提示),↑↓ 选、Enter/Tab 接受(插入 `NAME(` 并把光标移进括号)、Esc 关。
 * 只在公式(`=` 开头)且光标处于"函数名 token"时弹;不影响普通文本编辑。
 */
import { FUNCTION_NAMES, FUNCTION_SIGNATURES } from '../formula/builtin/functions'

export class FormulaAutocomplete {
  private box: HTMLElement | null = null
  private items: string[] = []
  private active = 0
  private tokenStart = 0 // 当前 token 在 value 中的起始下标

  constructor(private ta: HTMLTextAreaElement, private names: string[] = FUNCTION_NAMES) {}

  /** 输入变化后调用:重算建议并显示/隐藏。 */
  update(): void {
    const tok = this.currentToken()
    if (!tok) return this.close()
    const upper = tok.text.toUpperCase()
    const matches = this.names.filter((n) => n.startsWith(upper))
    // 仅剩一个且已完整输入 → 不打扰
    if (!matches.length || (matches.length === 1 && matches[0] === upper)) return this.close()
    this.tokenStart = tok.start
    this.items = matches.slice(0, 8)
    this.active = 0
    this.render()
  }

  isOpen(): boolean { return !!this.box }

  /** 列表打开时拦截导航键;返回 true = 已处理(编辑器应跳过自己的处理)。 */
  onKeyDown(e: KeyboardEvent): boolean {
    if (!this.box) return false
    if (e.key === 'ArrowDown') { this.active = (this.active + 1) % this.items.length; this.render(); return true }
    if (e.key === 'ArrowUp') { this.active = (this.active - 1 + this.items.length) % this.items.length; this.render(); return true }
    if (e.key === 'Enter' || e.key === 'Tab') { this.accept(this.items[this.active]); return true }
    if (e.key === 'Escape') { this.close(); return true }
    return false
  }

  close(): void {
    this.box?.remove()
    this.box = null
  }
  dispose(): void { this.close() }

  // ---------------- 内部 ----------------
  /** 光标前的"函数名 token":公式串里,光标紧邻一段字母,且其前是可起函数的位置(= ( , 运算符 空白 起始)。 */
  private currentToken(): { text: string; start: number } | null {
    const v = this.ta.value
    if (v[0] !== '=') return null
    const caret = this.ta.selectionStart ?? v.length
    let s = caret
    while (s > 0 && /[A-Za-z]/.test(v[s - 1])) s--
    if (s === caret) return null // 光标前不是字母
    const prev = s > 0 ? v[s - 1] : '='
    if (!'=(,+-*/^&<>% '.includes(prev) && s !== 1) return null // 不是函数起始上下文
    return { text: v.slice(s, caret), start: s }
  }

  private accept(name: string): void {
    const v = this.ta.value
    const caret = this.ta.selectionStart ?? v.length
    const next = v.slice(0, this.tokenStart) + name + '(' + v.slice(caret)
    const pos = this.tokenStart + name.length + 1
    this.ta.value = next
    this.ta.setSelectionRange(pos, pos)
    this.close()
    this.ta.dispatchEvent(new Event('input', { bubbles: true })) // 触发编辑器撑高/回显
    this.ta.focus()
  }

  private render(): void {
    if (typeof document === 'undefined') return
    if (!this.box) {
      this.box = document.createElement('div')
      this.box.className = 'ooxml-formula-ac'
      this.box.style.cssText = "position:fixed;z-index:10002;min-width:200px;max-width:360px;background:#fff;border:1px solid #d0d5dd;border-radius:6px;box-shadow:0 6px 24px rgba(0,0,0,.16);font:13px/1.5 -apple-system,'Segoe UI','Microsoft YaHei',monospace;overflow:hidden;max-height:240px;overflow-y:auto"
      document.body.appendChild(this.box)
    }
    this.box.innerHTML = this.items.map((n, i) => {
      const sig = FUNCTION_SIGNATURES[n] || `${n}(…)`
      const bg = i === this.active ? 'background:#eef4ff;' : ''
      return `<div data-i="${i}" style="padding:5px 10px;cursor:pointer;${bg}"><b style="color:#1a56db">${escapeHtml(n)}</b> <span style="color:#98a2b3;font-size:12px">${escapeHtml(sig)}</span></div>`
    }).join('')
    this.box.querySelectorAll('[data-i]').forEach((el) => {
      el.addEventListener('mousedown', (ev) => { ev.preventDefault(); this.accept(this.items[Number((el as HTMLElement).dataset.i)]) })
    })
    // 定位:编辑器下方
    const r = this.ta.getBoundingClientRect()
    const vh = window.innerHeight
    const below = vh - r.bottom > 250 || r.top < 250
    this.box.style.left = `${Math.min(r.left, window.innerWidth - 364)}px`
    if (below) { this.box.style.top = `${r.bottom + 2}px`; this.box.style.bottom = '' }
    else { this.box.style.bottom = `${vh - r.top + 2}px`; this.box.style.top = '' }
  }
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string)
}
