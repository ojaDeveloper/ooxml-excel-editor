/**
 * 数字格式编辑对话框(框架无关 DOM,三壳共用一份 —— 同 conditional-format / paste-config 模式)。1.11.0 新增。
 * 工具栏「数字格式」打开:分类(常规/数值/货币/百分比/日期/时间/文本/自定义)+ 选项 → 实时预览(复用 number-format 引擎)→ 确定即 setStyle({ numFmt })。
 */
import { formatValue } from '../format/number-format'

export interface NumberFormatDialogOptions {
  /** 当前活动格的值(预览用;无数值时用内置示例) */
  sampleValue: number | string | boolean | Date | null
  /** 当前 numFmt(预选) */
  currentCode: string
  date1904: boolean
  onApply: (code: string) => void
}

type Cat = 'general' | 'number' | 'currency' | 'percent' | 'date' | 'time' | 'text' | 'custom'
const CAT_LABEL: Record<Cat, string> = { general: '常规', number: '数值', currency: '货币', percent: '百分比', date: '日期', time: '时间', text: '文本', custom: '自定义' }
const DATE_PRESETS: Array<[string, string]> = [
  ['yyyy/m/d', '2026/4/1'], ['yyyy-mm-dd', '2026-04-01'], ['m月d日', '4月1日'], ['yyyy年m月d日', '2026年4月1日'], ['mm/dd/yy', '04/01/26'], ['yyyy/m/d h:mm', '2026/4/1 13:05'],
]
const TIME_PRESETS: Array<[string, string]> = [['h:mm', '13:05'], ['h:mm:ss', '13:05:09'], ['h:mm AM/PM', '1:05 PM'], ['[h]:mm:ss', '总时长']]
const CURRENCY_SYMBOLS: Array<[string, string]> = [['¥', '¥ 人民币'], ['$', '$ 美元'], ['€', '€ 欧元'], ['£', '£ 英镑'], ['', '无符号']]

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string)
}

/** 由分类 + 选项构造 numFmt 代码。 */
function buildCode(cat: Cat, o: { decimals: number; thousands: boolean; negRed: boolean; symbol: string; preset: string }): string {
  const dec = o.decimals > 0 ? '.' + '0'.repeat(o.decimals) : ''
  switch (cat) {
    case 'general': return 'General'
    case 'text': return '@'
    case 'number': {
      const body = (o.thousands ? '#,##0' : '0') + dec
      return o.negRed ? `${body};[Red]-${body}` : body
    }
    case 'currency': {
      const body = `${o.symbol}#,##0${dec}`
      return o.negRed ? `${body};[Red]-${body}` : body
    }
    case 'percent': return '0' + dec + '%'
    case 'date':
    case 'time': return o.preset
    case 'custom': return o.preset || 'General'
  }
}

function catOfCode(code: string): Cat {
  if (!code || code === 'General') return 'general'
  if (code === '@') return 'text'
  if (/%/.test(code)) return 'percent'
  if (/[¥$€£]/.test(code)) return 'currency'
  if (/[yhsm]|AM\/PM|\[h\]/i.test(code) && /[ymdhs:]/.test(code) && !/#|0\.0/.test(code.replace(/\[.*?\]/g, ''))) {
    return /[hs]|AM\/PM|\[h\]/i.test(code) && !/[ymd]/.test(code) ? 'time' : 'date'
  }
  if (/^[#0,.]+$/.test(code.split(';')[0])) return 'number'
  return 'custom'
}

const BTN = 'height:30px;padding:0 14px;border-radius:5px;cursor:pointer;font:inherit'
const PRIMARY = `${BTN};border:1px solid #1b7f4d;background:#21a366;color:#fff`
const PLAIN = `${BTN};border:1px solid #d0d5dd;background:#fff;color:#1f2329`

export class NumberFormatDialogHost {
  private el: HTMLElement | null = null
  private cleanup: (() => void) | null = null
  private cat: Cat = 'general'
  private opt = { decimals: 2, thousands: true, negRed: false, symbol: '¥', preset: 'yyyy/m/d' }
  private sample: NumberFormatDialogOptions['sampleValue'] = 1234.567
  private date1904 = false
  private onApply: (code: string) => void = () => {}

  show(opts: NumberFormatDialogOptions): void {
    if (typeof document === 'undefined') return
    this.close()
    this.sample = opts.sampleValue
    this.date1904 = opts.date1904
    this.onApply = opts.onApply
    this.cat = catOfCode(opts.currentCode)
    // 从当前代码尽量回填选项
    const dm = /\.0+/.exec(opts.currentCode)
    this.opt.decimals = dm ? dm[0].length - 1 : (this.cat === 'general' ? 2 : 0)
    this.opt.thousands = /#,##0/.test(opts.currentCode)
    this.opt.negRed = /\[Red\]/i.test(opts.currentCode)
    const sym = /[¥$€£]/.exec(opts.currentCode)
    if (sym) this.opt.symbol = sym[0]
    if (this.cat === 'date' || this.cat === 'time' || this.cat === 'custom') this.opt.preset = opts.currentCode

    const mask = document.createElement('div')
    mask.className = 'ooxml-numfmt-mask'
    mask.style.cssText = 'position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.32);display:flex;align-items:center;justify-content:center;'
    const card = document.createElement('div')
    card.className = 'ooxml-numfmt-card'
    card.style.cssText = "width:460px;max-width:94vw;max-height:90vh;overflow:auto;background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.25);font:13px/1.5 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;color:#1f2329"
    mask.appendChild(card)
    document.body.appendChild(mask)
    this.el = mask
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') this.close() }
    mask.addEventListener('mousedown', (e) => { if (e.target === mask) this.close() })
    document.addEventListener('keydown', onKey, true)
    this.cleanup = () => document.removeEventListener('keydown', onKey, true)
    this.render(card)
  }

  private currentCode(): string {
    return buildCode(this.cat, this.opt)
  }

  private render(card: HTMLElement): void {
    const code = this.currentCode()
    const preview = formatValue(this.previewValue(), code, this.date1904).text
    const cats = (Object.keys(CAT_LABEL) as Cat[]).map((c) =>
      `<button data-cat="${c}" style="${this.cat === c ? PRIMARY : PLAIN};height:28px;padding:0 10px;margin:0 6px 6px 0">${CAT_LABEL[c]}</button>`).join('')
    card.innerHTML = [
      '<div style="display:flex;align-items:center;justify-content:space-between;padding:13px 16px;border-bottom:1px solid #eef0f2;font-weight:600;position:sticky;top:0;background:#fff">',
      '<span>设置单元格格式</span><button data-close style="border:0;background:none;font-size:20px;line-height:1;color:#8a8f98;cursor:pointer">×</button></div>',
      '<div style="padding:14px 16px">',
      `<div style="margin-bottom:10px">${cats}</div>`,
      `<div data-opts style="min-height:40px"></div>`,
      `<div style="margin-top:12px;display:flex;align-items:center;gap:8px"><span style="flex:0 0 64px;color:#475467">格式代码</span><input data-code value="${escapeHtml(code)}" style="${PLAIN};height:30px;flex:1;font-family:Consolas,monospace"></div>`,
      `<div style="margin-top:10px;padding:10px 12px;background:#f7f8fa;border:1px solid #eef0f2;border-radius:6px"><span style="color:#98a2b3">预览:</span> <b data-preview style="font-size:15px">${escapeHtml(preview)}</b></div>`,
      '</div>',
      '<div style="display:flex;justify-content:flex-end;gap:8px;padding:12px 16px;background:#fafbfc;border-top:1px solid #eef0f2;position:sticky;bottom:0">',
      `<button data-cancel style="${PLAIN}">取消</button><button data-ok style="${PRIMARY}">确定</button>`,
      '</div>',
    ].join('')
    this.renderOpts(card.querySelector('[data-opts]') as HTMLElement)

    card.querySelectorAll('[data-cat]').forEach((b) => b.addEventListener('click', () => {
      this.cat = (b as HTMLElement).dataset.cat as Cat
      if (this.cat === 'date') this.opt.preset = DATE_PRESETS[0][0]
      if (this.cat === 'time') this.opt.preset = TIME_PRESETS[0][0]
      this.render(card)
    }))
    card.querySelector('[data-close]')?.addEventListener('click', () => this.close())
    card.querySelector('[data-cancel]')?.addEventListener('click', () => this.close())
    card.querySelector('[data-ok]')?.addEventListener('click', () => {
      const codeInput = (card.querySelector('[data-code]') as HTMLInputElement).value.trim()
      this.onApply(codeInput || 'General')
      this.close()
    })
    // 手改格式代码 → 切到自定义并刷新预览
    const codeInput = card.querySelector('[data-code]') as HTMLInputElement
    codeInput.addEventListener('input', () => {
      ;(card.querySelector('[data-preview]') as HTMLElement).textContent = formatValue(this.previewValue(), codeInput.value, this.date1904).text
    })
  }

  private renderOpts(host: HTMLElement): void {
    const row = (label: string, inner: string) => `<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px"><span style="flex:0 0 64px;color:#475467">${label}</span>${inner}</div>`
    const parts: string[] = []
    if (this.cat === 'number' || this.cat === 'currency' || this.cat === 'percent') {
      parts.push(row('小数位数', `<input data-o="decimals" type="number" min="0" max="10" value="${this.opt.decimals}" style="${PLAIN};height:28px;width:70px">`))
    }
    if (this.cat === 'number' || this.cat === 'currency') {
      parts.push(row('选项', `<label style="cursor:pointer"><input type="checkbox" data-o="thousands"${this.opt.thousands ? ' checked' : ''}> 千分位</label> <label style="cursor:pointer;margin-left:14px"><input type="checkbox" data-o="negRed"${this.opt.negRed ? ' checked' : ''}> 负数红色</label>`))
    }
    if (this.cat === 'currency') {
      parts.push(row('货币符号', `<select data-o="symbol" style="${PLAIN};height:30px">${CURRENCY_SYMBOLS.map(([v, t]) => `<option value="${v}"${v === this.opt.symbol ? ' selected' : ''}>${escapeHtml(t)}</option>`).join('')}</select>`))
    }
    if (this.cat === 'date' || this.cat === 'time') {
      const list = this.cat === 'date' ? DATE_PRESETS : TIME_PRESETS
      parts.push(`<div style="display:flex;flex-direction:column;gap:4px">${list.map(([v, t]) => `<label style="cursor:pointer;padding:5px 8px;border:1px solid ${v === this.opt.preset ? '#21a366' : '#eef0f2'};border-radius:5px;display:flex;justify-content:space-between"><span><input type="radio" name="nfp" data-o="preset" value="${escapeHtml(v)}"${v === this.opt.preset ? ' checked' : ''}> ${escapeHtml(v)}</span><span style="color:#98a2b3">${escapeHtml(t)}</span></label>`).join('')}</div>`)
    }
    if (this.cat === 'general') parts.push('<div style="color:#98a2b3">不设特定格式,按内容自动显示。</div>')
    if (this.cat === 'text') parts.push('<div style="color:#98a2b3">按文本显示,数字也不参与计算格式。</div>')
    if (this.cat === 'custom') parts.push('<div style="color:#98a2b3">在下方「格式代码」里直接输入自定义格式(如 <code>0.00"元"</code>、<code>#,##0;[Red]-#,##0</code>)。</div>')
    host.innerHTML = parts.join('')
    host.querySelectorAll('[data-o]').forEach((el) => el.addEventListener('change', () => {
      const t = el as HTMLInputElement | HTMLSelectElement
      const k = t.dataset.o as keyof typeof this.opt
      if (k === 'thousands' || k === 'negRed') (this.opt[k] as boolean) = (t as HTMLInputElement).checked
      else if (k === 'decimals') this.opt.decimals = Math.max(0, Math.min(10, Number(t.value) || 0))
      else (this.opt[k] as string) = t.value
      // 重渲整卡(代码 + 预览 + 选项联动)
      const card = host.closest('.ooxml-numfmt-card') as HTMLElement
      if (card) this.render(card)
    }))
  }

  /** 预览取值:活动格有数值/日期就用它,否则按分类给个内置示例。 */
  private previewValue(): number | string | boolean | Date | null {
    const v = this.sample
    if (this.cat === 'date' || this.cat === 'time') return v instanceof Date ? v : new Date(Date.UTC(2026, 3, 1, 13, 5, 9))
    if (this.cat === 'text') return typeof v === 'string' ? v : String(v ?? '示例文本')
    if (typeof v === 'number') return v
    return 1234.567
  }

  close(): void {
    this.cleanup?.()
    this.cleanup = null
    this.el?.remove()
    this.el = null
  }
  dispose(): void { this.close() }
}
