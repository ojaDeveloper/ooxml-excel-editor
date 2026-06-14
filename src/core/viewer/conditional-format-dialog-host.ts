/**
 * 条件格式管理对话框(框架无关 DOM,同 pivot-dialog-host / paste-config-host 模式 —— 三壳共用一份,UI 天然 1:1)。
 * 工具栏「条件格式」打开:列出当前表所有规则(可删/可编辑)+ 新建(6 类:突出显示单元格 cellIs / 公式 expression /
 * 色阶 colorScale / 数据条 dataBar / 图标集 iconSet / 项目选取 top10)。确定时把整张规则集回调给控制器(单次撤销)。
 * 新建规则默认套到当前选区;编辑保留原区域。
 */
import type { ConditionalRule, MergeRange } from '../model/types'

export interface ConditionalDialogOptions {
  rules: ConditionalRule[]
  selection: MergeRange | null
  genId: () => string
  onApply: (rules: ConditionalRule[]) => void
}

const TYPE_LABEL: Record<string, string> = {
  cellIs: '突出显示单元格', expression: '公式', colorScale: '色阶', dataBar: '数据条', iconSet: '图标集', top10: '项目选取(前/后 N)', unsupported: '未支持',
}
const OPERATORS: Array<[string, string]> = [
  ['greaterThan', '大于'], ['lessThan', '小于'], ['between', '介于'], ['notBetween', '不介于'],
  ['equal', '等于'], ['notEqual', '不等于'], ['greaterThanOrEqual', '大于等于'], ['lessThanOrEqual', '小于等于'],
]
const ICON_SETS: Array<[string, string]> = [
  ['3TrafficLights1', '三色交通灯'], ['3Arrows', '三向箭头'], ['3Symbols', '三符号(圈)'], ['3Flags', '三色旗'],
  ['4Rating', '四等评级'], ['5Rating', '五等评级'], ['5Quarters', '五象限'],
]

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string)
}
function colLabel(col: number): string {
  let s = ''; let n = col
  do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1 } while (n >= 0)
  return s
}
const refOf = (r: MergeRange) => `${colLabel(r.left)}${r.top + 1}:${colLabel(r.right)}${r.bottom + 1}`
const rangesText = (rs: MergeRange[]) => (rs || []).map(refOf).join(', ') || '—'

/** 规则摘要(列表行右侧灰字)。 */
function summarize(r: ConditionalRule): string {
  switch (r.type) {
    case 'cellIs': return `${(OPERATORS.find((o) => o[0] === r.operator)?.[1]) || r.operator || ''} ${(r.formulae || []).join(' , ')}`
    case 'expression': return (r.formulae || [])[0] || ''
    case 'colorScale': return r.colorScale ? (r.colorScale.mid ? '三色' : '双色') : ''
    case 'dataBar': return r.dataBar?.color || ''
    case 'iconSet': return (ICON_SETS.find((i) => i[0] === r.iconSet?.name)?.[1]) || r.iconSet?.name || ''
    case 'top10': return r.top10 ? `${r.top10.bottom ? '后' : '前'} ${r.top10.rank}${r.top10.percent ? '%' : ' 项'}` : ''
    default: return ''
  }
}

const BTN = 'height:30px;padding:0 14px;border-radius:5px;cursor:pointer;font:inherit'
const PRIMARY = `${BTN};border:1px solid #1b7f4d;background:#21a366;color:#fff`
const PLAIN = `${BTN};border:1px solid #d0d5dd;background:#fff;color:#1f2329`

export class ConditionalFormatDialogHost {
  private el: HTMLElement | null = null
  private cleanup: (() => void) | null = null
  private rules: ConditionalRule[] = []
  private selection: MergeRange | null = null
  private genId: () => string = () => 'cf-u0'
  private onApply: (rules: ConditionalRule[]) => void = () => {}
  private editing: { index: number; draft: ConditionalRule } | null = null

  show(opts: ConditionalDialogOptions): void {
    if (typeof document === 'undefined') return
    this.close()
    this.rules = opts.rules.map((r) => ({ ...r }))
    this.selection = opts.selection
    this.genId = opts.genId
    this.onApply = opts.onApply
    this.editing = null
    const mask = document.createElement('div')
    mask.className = 'ooxml-cf-mask'
    mask.style.cssText = 'position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.32);display:flex;align-items:center;justify-content:center;'
    const card = document.createElement('div')
    card.className = 'ooxml-cf-card'
    card.style.cssText = "width:520px;max-width:94vw;max-height:90vh;overflow:auto;background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.25);font:13px/1.5 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;color:#1f2329;"
    mask.appendChild(card)
    document.body.appendChild(mask)
    this.el = mask
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') this.close() }
    mask.addEventListener('mousedown', (e) => { if (e.target === mask) this.close() })
    document.addEventListener('keydown', onKey, true)
    this.cleanup = () => document.removeEventListener('keydown', onKey, true)
    this.renderList(card)
  }

  // ---------------- 列表视图 ----------------
  private renderList(card: HTMLElement): void {
    const rows = this.rules.map((r, i) => [
      '<div style="display:flex;align-items:center;gap:10px;padding:9px 0;border-bottom:1px dashed #eef0f2">',
      `<span style="flex:0 0 96px;font-weight:500">${escapeHtml(TYPE_LABEL[r.type] || r.type)}</span>`,
      `<span style="flex:1;min-width:0;color:#667085;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(summarize(r))} · <span style="color:#98a2b3">${escapeHtml(rangesText(r.ranges))}</span></span>`,
      r.type === 'unsupported' ? '' : `<button data-edit="${i}" style="${PLAIN};height:26px;padding:0 10px">编辑</button>`,
      `<button data-del="${i}" style="${PLAIN};height:26px;padding:0 10px;color:#d92d20;border-color:#fda29b">删除</button>`,
      '</div>',
    ].join('')).join('')
    card.innerHTML = [
      '<div style="display:flex;align-items:center;justify-content:space-between;padding:13px 16px;border-bottom:1px solid #eef0f2;font-weight:600;position:sticky;top:0;background:#fff">',
      '<span>条件格式</span><button data-close style="border:0;background:none;font-size:20px;line-height:1;color:#8a8f98;cursor:pointer">×</button></div>',
      '<div style="padding:10px 16px">',
      `<div style="color:#98a2b3;font-size:12px;margin-bottom:6px">当前选区:${escapeHtml(this.selection ? refOf(this.selection) : '无(新建规则将套到全表 A1)')}</div>`,
      this.rules.length ? rows : '<div style="color:#98a2b3;padding:16px 0;text-align:center">暂无条件格式规则</div>',
      '<div style="margin-top:12px"><button data-add style="' + PLAIN + '">+ 新建规则</button></div>',
      '</div>',
      '<div style="display:flex;justify-content:flex-end;gap:8px;padding:12px 16px;background:#fafbfc;border-top:1px solid #eef0f2;position:sticky;bottom:0">',
      `<button data-cancel style="${PLAIN}">取消</button><button data-done style="${PRIMARY}">完成</button>`,
      '</div>',
    ].join('')
    card.querySelector('[data-close]')?.addEventListener('click', () => this.close())
    card.querySelector('[data-cancel]')?.addEventListener('click', () => this.close())
    card.querySelector('[data-done]')?.addEventListener('click', () => { this.onApply(this.rules); this.close() })
    card.querySelector('[data-add]')?.addEventListener('click', () => this.startEdit(card, -1))
    card.querySelectorAll('[data-edit]').forEach((b) => b.addEventListener('click', () => this.startEdit(card, Number((b as HTMLElement).dataset.edit))))
    card.querySelectorAll('[data-del]').forEach((b) => b.addEventListener('click', () => {
      this.rules.splice(Number((b as HTMLElement).dataset.del), 1)
      this.renderList(card)
    }))
  }

  // ---------------- 编辑视图 ----------------
  private startEdit(card: HTMLElement, index: number): void {
    const existing = index >= 0 ? this.rules[index] : null
    const ranges = existing ? existing.ranges : (this.selection ? [this.selection] : [{ top: 0, left: 0, bottom: 0, right: 0 }])
    const draft: ConditionalRule = existing
      ? { ...existing }
      : { id: this.genId(), origin: 'user', ranges, priority: 0, type: 'cellIs', operator: 'greaterThan', formulae: [''], style: { fill: { type: 'solid', fgColor: '#FFEB9C' }, font: { color: '#9C5700' } } }
    this.editing = { index, draft }
    this.renderEditor(card)
  }

  private renderEditor(card: HTMLElement): void {
    const d = this.editing!.draft
    const isNew = this.editing!.index < 0
    const typeSel = isNew
      ? `<select data-type style="${PLAIN};height:30px;width:180px">${['cellIs', 'colorScale', 'dataBar', 'iconSet', 'top10', 'expression'].map((t) => `<option value="${t}"${t === d.type ? ' selected' : ''}>${escapeHtml(TYPE_LABEL[t])}</option>`).join('')}</select>`
      : `<b>${escapeHtml(TYPE_LABEL[d.type] || d.type)}</b>`
    card.innerHTML = [
      '<div style="display:flex;align-items:center;justify-content:space-between;padding:13px 16px;border-bottom:1px solid #eef0f2;font-weight:600;position:sticky;top:0;background:#fff">',
      `<span>${isNew ? '新建规则' : '编辑规则'}</span><button data-close style="border:0;background:none;font-size:20px;line-height:1;color:#8a8f98;cursor:pointer">×</button></div>`,
      '<div style="padding:14px 16px">',
      `<div style="display:flex;align-items:center;gap:10px;margin-bottom:12px"><span style="flex:0 0 70px;color:#475467">规则类型</span>${typeSel}</div>`,
      `<div data-fields></div>`,
      `<div style="color:#98a2b3;font-size:12px;margin-top:12px">应用范围:${escapeHtml(rangesText(d.ranges))}</div>`,
      '</div>',
      '<div style="display:flex;justify-content:flex-end;gap:8px;padding:12px 16px;background:#fafbfc;border-top:1px solid #eef0f2;position:sticky;bottom:0">',
      `<button data-back style="${PLAIN}">返回</button><button data-save style="${PRIMARY}">保存</button>`,
      '</div>',
    ].join('')
    const fields = card.querySelector('[data-fields]') as HTMLElement
    this.renderFields(fields, d)
    card.querySelector('[data-type]')?.addEventListener('change', (e) => {
      this.editing!.draft = this.defaultForType((e.target as HTMLSelectElement).value, d.ranges, d.id ?? this.genId())
      this.renderEditor(card)
    })
    card.querySelector('[data-close]')?.addEventListener('click', () => this.close())
    card.querySelector('[data-back]')?.addEventListener('click', () => this.renderList(card))
    card.querySelector('[data-save]')?.addEventListener('click', () => {
      this.collectFields(fields, this.editing!.draft)
      const { index, draft } = this.editing!
      if (index >= 0) this.rules[index] = { ...draft, dirty: true }
      else this.rules.push(draft)
      this.editing = null
      this.renderList(card)
    })
  }

  private defaultForType(type: string, ranges: MergeRange[], id: string): ConditionalRule {
    const base = { id, origin: 'user' as const, ranges, priority: 0 }
    switch (type) {
      case 'colorScale': return { ...base, type: 'colorScale', colorScale: { min: '#F8696B', mid: '#FFEB84', max: '#63BE7B' } }
      case 'dataBar': return { ...base, type: 'dataBar', dataBar: { color: '#638EC6', gradient: true } }
      case 'iconSet': return { ...base, type: 'iconSet', iconSet: { name: '3TrafficLights1', reverse: false } }
      case 'top10': return { ...base, type: 'top10', top10: { rank: 10, percent: false, bottom: false }, style: { fill: { type: 'solid', fgColor: '#FFC7CE' }, font: { color: '#9C0006' } } }
      case 'expression': return { ...base, type: 'expression', formulae: [''], style: { fill: { type: 'solid', fgColor: '#FFEB9C' } } }
      default: return { ...base, type: 'cellIs', operator: 'greaterThan', formulae: [''], style: { fill: { type: 'solid', fgColor: '#FFEB9C' }, font: { color: '#9C5700' } } }
    }
  }

  /** 类型专属字段。 */
  private renderFields(host: HTMLElement, d: ConditionalRule): void {
    const row = (label: string, inner: string) => `<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px"><span style="flex:0 0 70px;color:#475467">${escapeHtml(label)}</span>${inner}</div>`
    const colorInput = (key: string, val: string) => `<input type="color" data-f="${key}" value="${escapeHtml(val || '#FFFFFF')}" style="width:42px;height:28px;padding:0;border:1px solid #d0d5dd;border-radius:4px;cursor:pointer">`
    const parts: string[] = []
    if (d.type === 'cellIs') {
      const op = d.operator || 'greaterThan'
      const two = op === 'between' || op === 'notBetween'
      parts.push(row('条件', `<select data-f="operator" style="${PLAIN};height:30px;width:140px">${OPERATORS.map(([v, t]) => `<option value="${v}"${v === op ? ' selected' : ''}>${escapeHtml(t)}</option>`).join('')}</select>`))
      parts.push(row('值', `<input data-f="v0" value="${escapeHtml((d.formulae || [])[0] ?? '')}" placeholder="数值/文本" style="${PLAIN};height:30px;width:110px"> ${two ? `<input data-f="v1" value="${escapeHtml((d.formulae || [])[1] ?? '')}" placeholder="到" style="${PLAIN};height:30px;width:110px">` : ''}`))
      parts.push(this.formatFields(d))
    } else if (d.type === 'expression') {
      parts.push(row('公式', `<input data-f="v0" value="${escapeHtml((d.formulae || [])[0] ?? '')}" placeholder="=$A1>100" style="${PLAIN};height:30px;flex:1">`))
      parts.push(this.formatFields(d))
    } else if (d.type === 'colorScale') {
      const cs = d.colorScale || { min: '#F8696B', max: '#63BE7B' }
      const three = !!cs.mid
      parts.push(row('色标数', `<label style="cursor:pointer"><input type="radio" name="csn" data-f="csn" value="2"${three ? '' : ' checked'}> 双色</label> <label style="cursor:pointer;margin-left:12px"><input type="radio" name="csn" data-f="csn" value="3"${three ? ' checked' : ''}> 三色</label>`))
      parts.push(row('最小', colorInput('csmin', cs.min)))
      if (three) parts.push(row('中间', colorInput('csmid', cs.mid || '#FFEB84')))
      parts.push(row('最大', colorInput('csmax', cs.max)))
    } else if (d.type === 'dataBar') {
      parts.push(row('条颜色', colorInput('dbcolor', d.dataBar?.color || '#638EC6')))
      parts.push(row('渐变', `<label style="cursor:pointer"><input type="checkbox" data-f="dbgrad"${d.dataBar?.gradient !== false ? ' checked' : ''}> 渐变填充</label>`))
    } else if (d.type === 'iconSet') {
      const name = d.iconSet?.name || '3TrafficLights1'
      parts.push(row('图标集', `<select data-f="iconset" style="${PLAIN};height:30px;width:160px">${ICON_SETS.map(([v, t]) => `<option value="${v}"${v === name ? ' selected' : ''}>${escapeHtml(t)}</option>`).join('')}</select>`))
      parts.push(row('反向', `<label style="cursor:pointer"><input type="checkbox" data-f="iconrev"${d.iconSet?.reverse ? ' checked' : ''}> 反转图标顺序</label>`))
    } else if (d.type === 'top10') {
      const t = d.top10 || { rank: 10, percent: false, bottom: false }
      parts.push(row('选取', `<label style="cursor:pointer"><input type="radio" name="t10b" data-f="t10top" value="top"${t.bottom ? '' : ' checked'}> 前</label> <label style="cursor:pointer;margin-left:12px"><input type="radio" name="t10b" data-f="t10top" value="bottom"${t.bottom ? ' checked' : ''}> 后</label> <input data-f="t10rank" type="number" min="1" value="${t.rank}" style="${PLAIN};height:30px;width:70px;margin-left:10px"> <label style="cursor:pointer;margin-left:8px"><input type="checkbox" data-f="t10pct"${t.percent ? ' checked' : ''}> 百分比</label>`))
      parts.push(this.formatFields(d))
    }
    host.innerHTML = parts.join('')
    // 色阶/top10 的单选切换需重渲(增减中间色标)
    host.querySelectorAll('[data-f="csn"]').forEach((el) => el.addEventListener('change', () => {
      this.collectFields(host, this.editing!.draft)
      const n = (host.querySelector('[data-f="csn"]:checked') as HTMLInputElement)?.value
      const cs = this.editing!.draft.colorScale || { min: '#F8696B', max: '#63BE7B' }
      this.editing!.draft.colorScale = n === '3' ? { min: cs.min, mid: cs.mid || '#FFEB84', max: cs.max } : { min: cs.min, max: cs.max }
      this.renderFields(host, this.editing!.draft)
    }))
  }

  /** 通用「格式」字段(填充色 + 字体色 + 加粗),cellIs/expression/top10 用。 */
  private formatFields(d: ConditionalRule): string {
    const fill = d.style?.fill && d.style.fill.type !== 'none' ? (d.style.fill.fgColor || '#FFEB9C') : ''
    const font = d.style?.font?.color || ''
    const bold = !!d.style?.font?.bold
    return [
      '<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">',
      '<span style="flex:0 0 70px;color:#475467">命中格式</span>',
      `<label style="cursor:pointer"><input type="checkbox" data-f="hasfill"${fill ? ' checked' : ''}> 填充</label> <input type="color" data-f="fillcolor" value="${escapeHtml(fill || '#FFEB9C')}" style="width:36px;height:26px;padding:0;border:1px solid #d0d5dd;border-radius:4px;cursor:pointer">`,
      `<label style="cursor:pointer;margin-left:8px"><input type="checkbox" data-f="hasfont"${font ? ' checked' : ''}> 字体色</label> <input type="color" data-f="fontcolor" value="${escapeHtml(font || '#9C0006')}" style="width:36px;height:26px;padding:0;border:1px solid #d0d5dd;border-radius:4px;cursor:pointer">`,
      `<label style="cursor:pointer;margin-left:8px"><input type="checkbox" data-f="bold"${bold ? ' checked' : ''}> 加粗</label>`,
      '</div>',
    ].join('')
  }

  /** 从 DOM 收集字段写回 draft。 */
  private collectFields(host: HTMLElement, d: ConditionalRule): void {
    const val = (k: string) => (host.querySelector(`[data-f="${k}"]`) as HTMLInputElement | HTMLSelectElement | null)?.value
    const checked = (k: string) => (host.querySelector(`[data-f="${k}"]`) as HTMLInputElement | null)?.checked
    if (d.type === 'cellIs') {
      d.operator = val('operator') || 'greaterThan'
      const two = d.operator === 'between' || d.operator === 'notBetween'
      d.formulae = two ? [val('v0') ?? '', val('v1') ?? ''] : [val('v0') ?? '']
      this.collectFormat(host, d)
    } else if (d.type === 'expression') {
      d.formulae = [val('v0') ?? '']
      this.collectFormat(host, d)
    } else if (d.type === 'colorScale') {
      const n = (host.querySelector('[data-f="csn"]:checked') as HTMLInputElement)?.value || (d.colorScale?.mid ? '3' : '2')
      const min = val('csmin') || '#F8696B', max = val('csmax') || '#63BE7B'
      d.colorScale = n === '3' ? { min, mid: val('csmid') || '#FFEB84', max } : { min, max }
    } else if (d.type === 'dataBar') {
      d.dataBar = { color: val('dbcolor') || '#638EC6', gradient: !!checked('dbgrad') }
    } else if (d.type === 'iconSet') {
      d.iconSet = { name: val('iconset') || '3TrafficLights1', reverse: !!checked('iconrev') }
    } else if (d.type === 'top10') {
      d.top10 = { rank: Math.max(1, Number(val('t10rank')) || 10), percent: !!checked('t10pct'), bottom: (host.querySelector('[data-f="t10top"]:checked') as HTMLInputElement)?.value === 'bottom' }
      this.collectFormat(host, d)
    }
  }

  private collectFormat(host: HTMLElement, d: ConditionalRule): void {
    const val = (k: string) => (host.querySelector(`[data-f="${k}"]`) as HTMLInputElement | null)?.value
    const checked = (k: string) => (host.querySelector(`[data-f="${k}"]`) as HTMLInputElement | null)?.checked
    const style: NonNullable<ConditionalRule['style']> = {}
    if (checked('hasfill')) style.fill = { type: 'solid', fgColor: val('fillcolor') || '#FFEB9C' }
    else style.fill = { type: 'none' }
    const font: Record<string, unknown> = {}
    if (checked('hasfont')) font.color = val('fontcolor') || '#9C0006'
    if (checked('bold')) font.bold = true
    if (Object.keys(font).length) style.font = font as NonNullable<ConditionalRule['style']>['font']
    d.style = style
  }

  close(): void {
    this.cleanup?.()
    this.cleanup = null
    this.el?.remove()
    this.el = null
    this.editing = null
  }
  dispose(): void { this.close() }
}
