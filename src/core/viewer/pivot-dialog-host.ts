/** 透视表字段选择对话框(框架无关 DOM),供 Vue/React/Vue2 共用同一入口。 */
import type { PivotFilterMode, PivotFilterRule, PivotSummary, PivotTableLayout as PivotPanelLayout, PivotValueRule } from '../model/types'

/** 四区大白话说明(hover 区域标题/方框显示的 native tooltip)。 */
const AREA_HELP: Record<string, string> = {
  filters: '筛选器:决定整张透视表只统计哪些数据。把字段拖进来后,点字段上的小按钮选「非空」或「多选某几个值」,表里就只算选中的部分;放进来默认是「全部」=不过滤(所以看不出变化,要点开选值才生效)。',
  columns: '列:把这个字段的每个不同值拆成一列(横向展开),和「行」交叉成二维汇总表。例:行放「物流商」、列放「账单月份」→ 行是每个物流商,每个月一列,格子里是该物流商该月的合计。',
  rows: '行:把这个字段的每个不同值列成一行(纵向分组),格子里是对应的汇总值。放 2 个及以上行字段时,外层分组可点行首的 − / + 折叠或展开下面的明细。',
  values: '值:要汇总的数字字段(放进来才会出数据)。点字段上的下拉可切换 求和 / 计数 / 平均 / 最大 / 最小;可以同时放多个值字段,各占一列。',
}
/** 筛选器/列空着时在方框里显示的一句话引导(行/值已经很直观,不加)。 */
const AREA_HINT: Record<string, string> = {
  filters: '放字段并选值 → 只统计选中的数据',
  columns: '放字段 → 按它的值横向拆成多列',
}

export interface PivotFieldOption {
  index: number
  label: string
  numeric: boolean
}

export interface PivotDialogOptions {
  rangeLabel: string
  defaultOutputCell: string
  onSubmit: (output: PivotOutputChoice) => void
}

export type PivotOutputChoice = { kind: 'current-sheet'; cell: string } | { kind: 'new-sheet' }

export interface PivotPanelOptions {
  fields: PivotFieldOption[]
  filterValues: Record<number, string[]>
  layout: PivotPanelLayout
  onChange: (layout: PivotPanelLayout) => void
}

export class PivotDialogHost {
  private el: HTMLElement | null = null
  private cleanup: (() => void) | null = null

  show(opts: PivotDialogOptions): void {
    if (typeof document === 'undefined') return
    this.close()
    const mask = document.createElement('div')
    mask.style.cssText = 'position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.32);display:flex;align-items:center;justify-content:center;'
    const card = document.createElement('div')
    card.style.cssText = "width:420px;max-width:92vw;background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.25);font:13px/1.5 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;color:#1f2329;overflow:hidden;"
    card.innerHTML = [
      '<div style="display:flex;align-items:center;justify-content:space-between;padding:13px 16px;border-bottom:1px solid #eef0f2;font-weight:600">',
      '<span>创建透视表</span><button data-close style="border:0;background:none;font-size:20px;line-height:1;color:#8a8f98;cursor:pointer">×</button></div>',
      '<div style="padding:14px 16px 8px">',
      `<div style="margin-bottom:10px;color:#667085">数据区域: <b style="color:#1f2329">${escapeHtml(opts.rangeLabel)}</b></div>`,
      '<div style="margin-bottom:8px;color:#667085">请选择透视表生成位置。字段布局将在右侧“数据透视表”面板中配置。</div>',
      '<label style="display:block;margin:12px 0 4px;color:#667085">生成位置</label>',
      '<label style="display:flex;align-items:center;gap:6px;margin:4px 0"><input data-output-current type="radio" name="pivot-output" checked> 现有工作表</label>',
      `<input data-cell value="${escapeHtml(opts.defaultOutputCell)}" style="width:100%;height:30px;border:1px solid #d0d5dd;border-radius:5px;padding:0 8px;box-sizing:border-box" />`,
      '<label style="display:flex;align-items:center;gap:6px;margin:8px 0 0"><input data-output-new type="radio" name="pivot-output"> 新建工作表</label>',
      '<div style="margin-top:10px;color:#98a2b3;font-size:12px">字段、筛选、折叠可在右侧面板调整;源数据改动后透视结果自动刷新。</div>',
      '</div>',
      '<div style="display:flex;gap:8px;justify-content:flex-end;padding:12px 16px;background:#fafbfc;border-top:1px solid #eef0f2">',
      '<button data-cancel style="height:30px;padding:0 14px;border:1px solid #d0d5dd;border-radius:5px;background:#fff;cursor:pointer">取消</button>',
      '<button data-ok style="height:30px;padding:0 14px;border:1px solid #1b7f4d;border-radius:5px;background:#21a366;color:#fff;cursor:pointer">创建</button>',
      '</div>',
    ].join('')
    mask.appendChild(card)
    document.body.appendChild(mask)
    this.el = mask

    const cellInput = card.querySelector('[data-cell]') as HTMLInputElement
    const currentRadio = card.querySelector('[data-output-current]') as HTMLInputElement
    const newRadio = card.querySelector('[data-output-new]') as HTMLInputElement

    const submit = () => {
      const output: PivotOutputChoice = newRadio.checked ? { kind: 'new-sheet' } : { kind: 'current-sheet', cell: cellInput.value.trim() }
      this.close()
      opts.onSubmit(output)
    }
    cellInput.addEventListener('focus', () => { currentRadio.checked = true })
    currentRadio.addEventListener('change', () => { if (currentRadio.checked) cellInput.focus() })
    newRadio.addEventListener('change', () => { if (newRadio.checked) cellInput.blur() })
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') this.close()
      if (e.key === 'Enter') submit()
    }
    card.querySelector('[data-close]')?.addEventListener('click', () => this.close())
    card.querySelector('[data-cancel]')?.addEventListener('click', () => this.close())
    card.querySelector('[data-ok]')?.addEventListener('click', submit)
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
  }
}

export class PivotFieldPanelHost {
  private el: HTMLElement | null = null
  private opts: PivotPanelOptions | null = null
  /** 当前正在编辑筛选值的字段 index(null = 未展开筛选明细面板) */
  private activeFilterField: number | null = null

  show(opts: PivotPanelOptions): void {
    if (typeof document === 'undefined') return
    this.close()
    this.opts = opts
    this.activeFilterField = null
    const panel = document.createElement('div')
    panel.style.cssText = "position:fixed;right:0;top:0;bottom:0;width:360px;z-index:9999;background:#e9eef4;border-left:1px solid #cfd6df;box-shadow:-4px 0 16px rgba(15,23,42,.12);font:13px/1.45 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;color:#111827;padding:18px 16px;box-sizing:border-box;overflow:auto;"
    document.body.appendChild(panel)
    this.el = panel
    this.render()
  }

  close(): void {
    this.el?.remove()
    this.el = null
    this.opts = null
    this.activeFilterField = null
  }

  dispose(): void {
    this.close()
  }

  private render(filter = ''): void {
    const panel = this.el
    const opts = this.opts
    if (!panel || !opts) return
    const q = filter.trim().toLowerCase()
    const visible = q ? opts.fields.filter((f) => f.label.toLowerCase().includes(q)) : opts.fields
    panel.innerHTML = [
      '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px">',
      '<div style="font-weight:600">数据透视表⌄</div><div style="display:flex;gap:12px;color:#374151"><span>⚙</span><span>⌖</span><button data-close style="border:0;background:none;font-size:18px;cursor:pointer;color:#374151">×</button></div></div>',
      '<details open><summary style="font-weight:600;cursor:pointer">字段列表</summary>',
      '<div style="margin:10px 0 6px">将字段添加至数据透视表区域</div>',
      `<input data-search value="${escapeHtml(filter)}" placeholder="  搜索字段" style="width:100%;height:32px;border:1px solid #c9d1db;border-radius:4px;background:#fff;box-sizing:border-box;margin-bottom:6px" />`,
      '<div style="height:270px;background:#fff;border:1px solid #c9d1db;border-radius:4px;overflow:auto;padding:8px 8px;box-sizing:border-box">',
      ...visible.map((field) => this.fieldRow(field, opts.layout)),
      '</div></details>',
      '<details open style="margin-top:18px"><summary style="font-weight:600;cursor:pointer">数据透视表区域</summary>',
      '<div style="margin:10px 0 8px">在下面区域中拖动字段</div>',
      '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">',
      this.area('筛选器', 'filters', opts),
      this.area('列', 'columns', opts),
      this.area('行', 'rows', opts),
      this.area('值', 'values', opts),
      '</div>',
      this.filterDetail(opts),
      '<div data-drop-remove style="margin-top:10px;height:30px;border:1px dashed #9aa4b2;border-radius:4px;display:flex;align-items:center;justify-content:center;color:#667085;background:#f7f9fb">拖到这里移出字段</div></details>',
    ].join('')
    panel.querySelector('[data-close]')?.addEventListener('click', () => this.close())
    const search = panel.querySelector('[data-search]') as HTMLInputElement | null
    search?.addEventListener('input', () => this.render(search.value))
    search?.focus()
    search?.setSelectionRange(search.value.length, search.value.length)
    panel.querySelectorAll<HTMLButtonElement>('[data-add]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const index = Number(btn.dataset.field)
        const area = btn.dataset.add as keyof PivotPanelLayout
        this.updateLayout((layout) => addToArea(layout, area, index), filter)
      })
    })
    // 字段列表复选框:勾选 = 加入(数值字段→值,其它→行,WPS 自动分区);取消 = 从所有区移出。
    panel.querySelectorAll<HTMLInputElement>('[data-toggle-field]').forEach((cb) => {
      cb.addEventListener('change', () => {
        const index = Number(cb.dataset.toggleField)
        const numeric = this.opts?.fields.find((f) => f.index === index)?.numeric ?? false
        this.updateLayout((layout) => cb.checked ? addToArea(layout, numeric ? 'values' : 'rows', index) : removeFromLayout(layout, index), filter)
      })
    })
    panel.querySelectorAll<HTMLButtonElement>('[data-remove]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const index = Number(btn.dataset.field)
        const area = btn.dataset.remove as keyof PivotPanelLayout
        this.updateLayout((layout) => area === 'filters'
          ? { ...layout, filters: layout.filters.filter((rule) => rule.field !== index) }
          : { ...layout, [area]: layout[area].filter((x) => x !== index) }, filter)
      })
    })
    panel.querySelectorAll<HTMLButtonElement>('[data-filter-edit]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const f = Number(btn.dataset.filterEdit)
        this.activeFilterField = this.activeFilterField === f ? null : f
        this.render(filter)
      })
    })
    panel.querySelector('[data-filter-close]')?.addEventListener('click', () => { this.activeFilterField = null; this.render(filter) })
    panel.querySelectorAll<HTMLInputElement>('[data-filter-mode]').forEach((radio) => {
      radio.addEventListener('change', () => {
        const field = this.activeFilterField
        if (field == null) return
        const mode = radio.dataset.filterMode as PivotFilterMode
        const values = this.opts?.filterValues[field] ?? []
        this.updateLayout((layout) => ({
          ...layout,
          filters: layout.filters.map((rule) => rule.field === field ? filterRuleForMode(field, mode, values, rule) : rule),
        }), filter)
      })
    })
    panel.querySelectorAll<HTMLInputElement>('[data-filter-check]').forEach((cb) => {
      cb.addEventListener('change', () => {
        const field = this.activeFilterField
        if (field == null) return
        const val = cb.dataset.filterCheck ?? ''
        const all = this.opts?.filterValues[field] ?? []
        this.updateLayout((layout) => ({
          ...layout,
          filters: layout.filters.map((rule) => {
            if (rule.field !== field) return rule
            const cur = new Set(rule.mode === 'include' ? (rule.values ?? all) : all)
            if (cb.checked) cur.add(val)
            else cur.delete(val)
            return { field, mode: 'include' as const, values: [...cur] }
          }),
        }), filter)
      })
    })
    panel.querySelector('[data-filter-all]')?.addEventListener('click', () => {
      const field = this.activeFilterField
      if (field == null) return
      const all = this.opts?.filterValues[field] ?? []
      this.updateLayout((layout) => ({ ...layout, filters: layout.filters.map((rule) => rule.field === field ? { field, mode: 'include' as const, values: all.slice() } : rule) }), filter)
    })
    panel.querySelector('[data-filter-none]')?.addEventListener('click', () => {
      const field = this.activeFilterField
      if (field == null) return
      this.updateLayout((layout) => ({ ...layout, filters: layout.filters.map((rule) => rule.field === field ? { field, mode: 'include' as const, values: [] } : rule) }), filter)
    })
    panel.querySelectorAll<HTMLSelectElement>('[data-summary-value]').forEach((select) => {
      select.addEventListener('change', () => {
        const field = Number(select.dataset.summaryValue)
        this.updateLayout((layout) => ({
          ...layout,
          values: layout.values.map((rule) => rule.field === field ? { ...rule, summary: select.value as PivotSummary } : rule),
        }), filter)
      })
    })
    panel.querySelectorAll<HTMLElement>('[data-drag-field]').forEach((el) => {
      el.addEventListener('dragstart', (e) => {
        e.dataTransfer?.setData('application/x-ooxml-pivot-field', JSON.stringify({ field: Number(el.dataset.dragField), from: el.dataset.dragFrom || '' }))
        if (e.dataTransfer) e.dataTransfer.effectAllowed = 'move'
      })
    })
    panel.querySelectorAll<HTMLElement>('[data-drop-area]').forEach((el) => {
      el.addEventListener('dragover', (e) => {
        e.preventDefault()
        el.style.outline = '2px solid #21a366'
      })
      el.addEventListener('dragleave', () => { el.style.outline = '' })
      el.addEventListener('drop', (e) => {
        e.preventDefault()
        el.style.outline = ''
        const raw = e.dataTransfer?.getData('application/x-ooxml-pivot-field')
        if (!raw) return
        const payload = JSON.parse(raw) as { field: number; from?: string }
        const area = el.dataset.dropArea as keyof PivotPanelLayout
        this.updateLayout((layout) => addToArea(layout, area, payload.field), filter)
      })
    })
    panel.querySelectorAll<HTMLElement>('[data-drop-remove]').forEach((el) => {
      el.addEventListener('dragover', (e) => {
        e.preventDefault()
        el.style.outline = '2px solid #ef4444'
      })
      el.addEventListener('dragleave', () => { el.style.outline = '' })
      el.addEventListener('drop', (e) => {
        e.preventDefault()
        el.style.outline = ''
        const raw = e.dataTransfer?.getData('application/x-ooxml-pivot-field')
        if (!raw) return
        const payload = JSON.parse(raw) as { field: number }
        this.updateLayout((layout) => removeFromLayout(layout, payload.field), filter)
      })
    })
  }

  private fieldRow(field: PivotFieldOption, layout: PivotPanelLayout): string {
    const checked = layout.filters.some((rule) => rule.field === field.index) || layout.columns.includes(field.index) || layout.rows.includes(field.index) || layout.values.some((rule) => rule.field === field.index)
    const addValue = field.numeric ? `<button data-add="values" data-field="${field.index}" title="加入值" style="border:0;background:none;cursor:pointer">Σ</button>` : ''
    return `<div data-drag-field="${field.index}" draggable="true" style="display:flex;align-items:center;gap:7px;min-height:30px;cursor:grab"><span style="color:#667085">⋮⋮</span><input type="checkbox" data-toggle-field="${field.index}" ${checked ? 'checked' : ''} title="勾选 = 加入透视表(数值→值,其它→行);取消 = 移出"> <span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(field.label)}</span><button data-add="filters" data-field="${field.index}" title="加入筛选器" style="border:0;background:none;cursor:pointer">筛</button><button data-add="columns" data-field="${field.index}" title="加入列" style="border:0;background:none;cursor:pointer">列</button><button data-add="rows" data-field="${field.index}" title="加入行" style="border:0;background:none;cursor:pointer">行</button>${addValue}</div>`
  }

  private area(title: string, area: keyof PivotPanelLayout, opts: PivotPanelOptions): string {
    const chips = opts.layout[area].map((index) => {
      const fieldIndex = typeof index === 'number' ? index : 'summary' in index ? (index as PivotValueRule).field : (index as PivotFilterRule).field
      const field = opts.fields.find((f) => f.index === fieldIndex)
      if (!field) return ''
      const label = area === 'values' ? `${summaryLabel((index as PivotValueRule).summary)}: ${field.label}` : field.label
      const filterBtn = area === 'filters' ? this.filterSummaryBtn(index as PivotFilterRule, opts.filterValues[field.index] ?? []) : ''
      const summarySelect = area === 'values' ? this.summarySelect(index as PivotValueRule) : ''
      return `<span data-drag-field="${field.index}" data-drag-from="${area}" draggable="true" style="display:inline-flex;align-items:center;gap:4px;max-width:100%;padding:3px 6px;margin:0 4px 4px 0;border:1px solid #b8c2cc;border-radius:3px;background:#f7f9fb;cursor:grab"><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(label)}</span>${filterBtn}${summarySelect}<button data-remove="${area}" data-field="${field.index}" style="border:0;background:none;cursor:pointer;color:#667085">×</button></span>`
    }).join('')
    const help = AREA_HELP[area]
    const hint = area === 'filters' || area === 'columns' ? `<div style="color:#9aa4ae;font-size:11px;padding:2px 0">${escapeHtml(AREA_HINT[area])}</div>` : ''
    return `<div title="${escapeHtml(help)}"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:5px"><span title="${escapeHtml(help)}" style="cursor:help">${escapeHtml(title)} <span style="color:#9aa4ae;font-size:11px">ⓘ</span></span><span>＋</span></div><div data-drop-area="${area}" title="${escapeHtml(help)}" style="height:88px;background:#fff;border:1px solid #c9d1db;border-radius:4px;padding:6px;box-sizing:border-box;overflow:auto">${chips || hint}</div></div>`
  }

  private updateLayout(update: (layout: PivotPanelLayout) => PivotPanelLayout, filter: string): void {
    if (!this.opts) return
    const next = update(this.opts.layout)
    this.opts.layout = next
    this.opts.onChange(next)
    this.render(filter)
  }

  /** 筛选 chip 上的小按钮:显示当前筛选状态(全部/非空/N 项),点开底部明细面板编辑。 */
  private filterSummaryBtn(rule: PivotFilterRule, values: string[]): string {
    const text = rule.mode === 'all' ? '全部'
      : rule.mode === 'non-empty' ? '非空'
      : rule.mode === 'equals' ? (rule.value ?? '')
      : `${rule.values?.length ?? values.length}/${values.length} 项`
    const active = this.activeFilterField === rule.field
    return `<button data-filter-edit="${rule.field}" title="设置筛选值" style="border:1px solid ${active ? '#21a366' : '#c9d1db'};background:#fff;border-radius:3px;height:22px;max-width:84px;padding:0 6px;cursor:pointer;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(text)} ▾</button>`
  }

  /** 底部筛选明细面板:全部/非空/多选 三模式 + 多选时的值复选框列表(WPS 风格勾选筛选)。 */
  private filterDetail(opts: PivotPanelOptions): string {
    const field = this.activeFilterField
    if (field == null) return ''
    const rule = opts.layout.filters.find((r) => r.field === field)
    if (!rule) return ''
    const fieldLabel = opts.fields.find((f) => f.index === field)?.label ?? `字段${field + 1}`
    const values = opts.filterValues[field] ?? []
    const isInclude = rule.mode === 'include'
    const selected = new Set(isInclude ? (rule.values ?? values) : values)
    const radio = (mode: string, label: string) => `<label style="display:inline-flex;align-items:center;gap:4px;margin-right:12px"><input type="radio" name="pivot-filter-mode" data-filter-mode="${mode}" ${rule.mode === mode ? 'checked' : ''}> ${label}</label>`
    const checks = values.map((v) => `<label style="display:flex;align-items:center;gap:6px;min-height:24px"><input type="checkbox" data-filter-check="${escapeHtml(v)}" ${selected.has(v) ? 'checked' : ''} ${isInclude ? '' : 'disabled'}> <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(v)}</span></label>`).join('')
    return [
      '<div style="margin-top:12px;border:1px solid #c9d1db;border-radius:4px;background:#fff;padding:8px">',
      `<div style="font-weight:600;margin-bottom:6px">筛选「${escapeHtml(fieldLabel)}」<button data-filter-close style="float:right;border:0;background:none;cursor:pointer;color:#667085">×</button></div>`,
      `<div style="margin-bottom:6px">${radio('all', '全部')}${radio('non-empty', '非空')}${radio('include', '多选')}</div>`,
      isInclude ? '<div style="margin-bottom:4px"><button data-filter-all style="border:0;background:none;color:#21a366;cursor:pointer;padding:0">全选</button> · <button data-filter-none style="border:0;background:none;color:#21a366;cursor:pointer;padding:0">清空</button></div>' : '',
      `<div style="max-height:140px;overflow:auto">${checks}</div>`,
      '</div>',
    ].join('')
  }

  private summarySelect(rule: PivotValueRule): string {
    const options: PivotSummary[] = ['sum', 'count', 'avg', 'max', 'min']
    return `<select data-summary-value="${rule.field}" style="max-width:78px;height:22px;border:1px solid #c9d1db;border-radius:3px;background:#fff">${options.map((value) => `<option value="${value}" ${value === rule.summary ? 'selected' : ''}>${summaryLabel(value)}</option>`).join('')}</select>`
  }
}

function addToArea(layout: PivotPanelLayout, area: keyof PivotPanelLayout, index: number): PivotPanelLayout {
  const next = removeFromLayout(layout, index)
  if (area === 'filters') next.filters = [...next.filters, { field: index, mode: 'all' }]
  else if (area === 'values') next.values = [...next.values, { field: index, summary: 'sum' }]
  else next[area] = [...next[area], index]
  return next
}

function removeFromLayout(layout: PivotPanelLayout, index: number): PivotPanelLayout {
  return {
    filters: layout.filters.filter((rule) => rule.field !== index),
    columns: layout.columns.filter((x) => x !== index),
    rows: layout.rows.filter((x) => x !== index),
    values: layout.values.filter((rule) => rule.field !== index),
  }
}

function summaryLabel(summary: PivotSummary): string {
  return { sum: '求和项', count: '计数项', avg: '平均值', max: '最大值', min: '最小值' }[summary]
}

/** 切换筛选模式时构造新规则:进入"多选"默认全选(沿用已有 include 选择),非空/全部清空值。 */
function filterRuleForMode(field: number, mode: PivotFilterMode, values: string[], prev: PivotFilterRule): PivotFilterRule {
  if (mode === 'include') return { field, mode: 'include', values: prev.mode === 'include' ? (prev.values ?? values.slice()) : values.slice() }
  if (mode === 'non-empty') return { field, mode: 'non-empty' }
  return { field, mode: 'all' }
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]!))
}
