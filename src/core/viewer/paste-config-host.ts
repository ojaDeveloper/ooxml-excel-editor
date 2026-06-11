/**
 * 粘贴行为配置面板(框架无关 DOM,同 pivot-dialog-host / context-menu 模式 —— 三壳共用一份实现,UI 天然 1:1)。
 * 工具栏「配置 ▾ → 粘贴行为」打开;列出 PasteBehavior 各项下拉自由定制 + 两个快捷预设;应用即调 setPasteBehavior。
 */
import { type PasteBehavior, DEFAULT_PASTE_BEHAVIOR, PASTE_PRESET_VALUES_ONLY, resolvePasteBehavior } from '../edit/paste-behavior'

export interface PasteConfigDialogOptions {
  current: PasteBehavior
  onSubmit: (cfg: PasteBehavior) => void
}

/** 每项:key + 中文标签 + 说明 + 选项[值,文案] */
const FIELDS: Array<{ key: keyof PasteBehavior; label: string; hint: string; opts: Array<[string, string]> }> = [
  { key: 'cellStyle', label: '字体/对齐/换行/边框/数字格式', hint: '覆盖=只留源;合并=源没写的留目标;不粘=保留目标', opts: [['overwrite', '覆盖'], ['merge', '合并'], ['skip', '不粘(留目标)']] },
  { key: 'fill', label: '填充底色', hint: '覆盖=源没写则无填充;合并=源没写则留目标底色;不粘=保留目标', opts: [['overwrite', '覆盖'], ['merge', '合并'], ['skip', '不粘(留目标)']] },
  { key: 'rowHeight', label: '行高', hint: '逐行,只影响被粘的行', opts: [['source', '搬源'], ['keep', '不动']] },
  { key: 'colWidth', label: '列宽', hint: '整列共享,搬到中间会动上方表头', opts: [['firstRowOnly', '仅首行搬源'], ['source', '总搬源'], ['keep', '不动']] },
  { key: 'targetMerges', label: '目标原有合并(粘贴区内)', hint: '不清的话旧合并会吞列致数据错位', opts: [['clear', '清掉'], ['keep', '保留']] },
  { key: 'sourceMerges', label: '源自带的合并', hint: '应用源里的合并区', opts: [['apply', '应用'], ['skip', '不应用']] },
  { key: 'images', label: '图片(内嵌/浮动)', hint: '把源的图落到目标格', opts: [['apply', '落格'], ['skip', '不粘']] },
]

function escapeHtml(s: string): string {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string)
}

export class PasteConfigDialogHost {
  private el: HTMLElement | null = null
  private cleanup: (() => void) | null = null

  show(opts: PasteConfigDialogOptions): void {
    if (typeof document === 'undefined') return
    this.close()
    const cur = resolvePasteBehavior(opts.current)
    const mask = document.createElement('div')
    mask.className = 'ooxml-paste-config-mask'
    mask.style.cssText = 'position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.32);display:flex;align-items:center;justify-content:center;'
    const card = document.createElement('div')
    card.style.cssText = "width:460px;max-width:94vw;max-height:90vh;overflow:auto;background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.25);font:13px/1.5 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;color:#1f2329;"

    const fieldRow = (f: (typeof FIELDS)[number]) => {
      const cur1 = String(cur[f.key])
      const options = f.opts.map(([v, t]) => `<option value="${v}"${v === cur1 ? ' selected' : ''}>${escapeHtml(t)}</option>`).join('')
      return [
        '<div style="display:flex;align-items:center;justify-content:space-between;gap:12px;padding:7px 0;border-bottom:1px dashed #eef0f2">',
        `<div style="min-width:0"><div style="font-weight:500">${escapeHtml(f.label)}</div><div style="color:#98a2b3;font-size:12px">${escapeHtml(f.hint)}</div></div>`,
        `<select data-key="${f.key}" style="flex:0 0 132px;height:30px;border:1px solid #d0d5dd;border-radius:5px;padding:0 6px;background:#fff;cursor:pointer">${options}</select>`,
        '</div>',
      ].join('')
    }

    card.innerHTML = [
      '<div style="display:flex;align-items:center;justify-content:space-between;padding:13px 16px;border-bottom:1px solid #eef0f2;font-weight:600;position:sticky;top:0;background:#fff">',
      '<span>粘贴行为配置</span><button data-close style="border:0;background:none;font-size:20px;line-height:1;color:#8a8f98;cursor:pointer">×</button></div>',
      '<div style="padding:10px 16px">',
      '<div style="display:flex;gap:8px;margin-bottom:8px">',
      '<button data-preset-overwrite style="flex:1;height:30px;border:1px solid #d0d5dd;border-radius:5px;background:#f5f7fa;cursor:pointer">覆盖式 1:1(默认)</button>',
      '<button data-preset-values style="flex:1;height:30px;border:1px solid #d0d5dd;border-radius:5px;background:#f5f7fa;cursor:pointer">仅值(保留目标)</button>',
      '</div>',
      '<div style="color:#98a2b3;font-size:12px;margin-bottom:4px">逐项自由定制(列宽默认「仅首行搬源」:粘到首行=新表头取源宽,粘到中间=不动上方表头)。</div>',
      FIELDS.map(fieldRow).join(''),
      '</div>',
      '<div style="display:flex;gap:8px;justify-content:flex-end;padding:12px 16px;background:#fafbfc;border-top:1px solid #eef0f2;position:sticky;bottom:0">',
      '<button data-reset style="height:30px;padding:0 12px;border:1px solid #d0d5dd;border-radius:5px;background:#fff;cursor:pointer">恢复默认</button>',
      '<button data-cancel style="height:30px;padding:0 14px;border:1px solid #d0d5dd;border-radius:5px;background:#fff;cursor:pointer">取消</button>',
      '<button data-ok style="height:30px;padding:0 14px;border:1px solid #1b7f4d;border-radius:5px;background:#21a366;color:#fff;cursor:pointer">应用</button>',
      '</div>',
    ].join('')
    mask.appendChild(card)
    document.body.appendChild(mask)
    this.el = mask

    const selects = () => Array.from(card.querySelectorAll('select[data-key]')) as HTMLSelectElement[]
    const setAll = (cfg: PasteBehavior) => { for (const s of selects()) s.value = String(cfg[s.dataset.key as keyof PasteBehavior]) }
    const collect = (): PasteBehavior => {
      const out = {} as Record<string, string>
      for (const s of selects()) out[s.dataset.key as string] = s.value
      return resolvePasteBehavior(out as Partial<PasteBehavior>)
    }
    const submit = () => { const cfg = collect(); this.close(); opts.onSubmit(cfg) }

    card.querySelector('[data-preset-overwrite]')?.addEventListener('click', () => setAll(DEFAULT_PASTE_BEHAVIOR))
    card.querySelector('[data-preset-values]')?.addEventListener('click', () => setAll(PASTE_PRESET_VALUES_ONLY))
    card.querySelector('[data-reset]')?.addEventListener('click', () => setAll(DEFAULT_PASTE_BEHAVIOR))
    card.querySelector('[data-close]')?.addEventListener('click', () => this.close())
    card.querySelector('[data-cancel]')?.addEventListener('click', () => this.close())
    card.querySelector('[data-ok]')?.addEventListener('click', submit)
    mask.addEventListener('mousedown', (e) => { if (e.target === mask) this.close() })
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') this.close() }
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
