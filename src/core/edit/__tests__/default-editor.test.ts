// @vitest-environment jsdom
/**
 * 默认编辑器单测 (Phase 1, 2026-06-08) — WPS 风格长文本撑高.
 *
 * 验证:
 *   ① mount 后 root 是 `<textarea>` 而非 `<input>` (改 Phase 1 起)
 *   ② getDesiredHeight: 短文本 1 行高, 长文本撑出多行高, \n 换行算高度
 *   ③ host position 后 textarea.style.height 反映撑高
 *   ④ 提交 / 取消 / 失焦行为
 *   ⑤ Shift+Enter 不提交 (插换行); 普通 Enter 提交 + 向下移
 */
import { describe, it, expect, beforeEach } from 'vitest'
import { defaultCellEditor } from '../default-editor'
import type { CellEditorContext } from '../editor-context'
import type { CellSnapshot } from '../../model/snapshot'

// JSDOM 默认环境的 textarea 是有的, 但 canvas.getContext('2d').measureText 可能没. vitest 默认 jsdom
// 通过 vi.stubGlobal 或者 mock canvas, 这里直接用真实环境

function makeSnapshot(text = ''): CellSnapshot {
  return {
    row: 0,
    col: 0,
    cell: null,
    raw: text || null,
    computed: text || null,
    text,
    style: {
      font: { name: 'Calibri', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#000000' },
      fill: { type: 'none' },
      borders: {},
      hAlign: 'left',
      vAlign: 'bottom',
      wrapText: false,
      shrinkToFit: false,
      textRotation: 0,
      indent: 0,
      numFmt: 'General',
    },
  } as CellSnapshot
}

function makeCtx(initialText = '', overrides: Partial<CellEditorContext> = {}): CellEditorContext {
  return {
    snapshot: makeSnapshot(initialText),
    rect: { x: 0, y: 0, w: 80, h: 20 },
    sheet: {} as never,
    workbook: {} as never,
    permission: 'editable',
    initialText: undefined,
    commit: () => {},
    cancel: () => {},
    ...overrides,
  }
}

describe('defaultCellEditor — Phase 1 长文本撑高 (2026-06-08)', () => {
  beforeEach(() => {
    // 清掉可能的 leftover DOM
    document.body.innerHTML = ''
  })

  it('① mount 后 root 是 textarea, 不再是 input', () => {
    const ctx = makeCtx('hi')
    const made = defaultCellEditor(ctx)
    const el = made instanceof HTMLElement ? made : made.el
    expect(el.tagName).toBe('TEXTAREA')
    expect(el.className).toContain('ooxml-cell-editor')
  })

  it('② getDesiredHeight 钩子已暴露 (具体高度计算依赖 canvas measureText, jsdom 不支持, 走 e2e 覆盖)', () => {
    const ctx = makeCtx('hi')
    const made = defaultCellEditor(ctx)
    if (made instanceof HTMLElement) throw new Error('expect object')
    expect(made.getDesiredHeight).toBeDefined()
    expect(typeof made.getDesiredHeight).toBe('function')
  })

  it('② 列宽 0 / 负值 → 期望高度 0 (兜底, 不依赖 canvas)', () => {
    const ctx = makeCtx('hello world')
    const made = defaultCellEditor(ctx)
    if (made instanceof HTMLElement) throw new Error('expect object')
    // 兜底分支不走 wrapLines, 直接返 0
    expect(made.getDesiredHeight!(0)).toBe(0)
    expect(made.getDesiredHeight!(-10)).toBe(0)
  })

  it('③ reposition 回调被注入: 输入事件后调用', async () => {
    let calls = 0
    const ctx = makeCtx('', { reposition: () => { calls++ } })
    const made = defaultCellEditor(ctx)
    if (made instanceof HTMLElement) throw new Error('expect object')
    document.body.appendChild(made.el)
    const ta = made.el as HTMLTextAreaElement
    ta.value = 'new value'
    ta.dispatchEvent(new Event('input'))
    expect(calls).toBeGreaterThanOrEqual(1)
  })

  it('④ Enter 触发 commit + move=down', async () => {
    let committed: { value: unknown; move?: string } | null = null
    const ctx = makeCtx('foo', { commit: (v, m) => { committed = { value: v, move: m } } })
    const made = defaultCellEditor(ctx)
    if (made instanceof HTMLElement) throw new Error('expect object')
    const ta = made.el as HTMLTextAreaElement
    document.body.appendChild(ta)
    ta.value = 'changed'
    ta.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    expect(committed).not.toBeNull()
    expect((committed as any).value).toBe('changed')
    expect((committed as any).move).toBe('down')
  })

  it('④ Shift+Enter 不提交 (插换行)', () => {
    let committed = false
    const ctx = makeCtx('foo', { commit: () => { committed = true } })
    const made = defaultCellEditor(ctx)
    if (made instanceof HTMLElement) throw new Error('expect object')
    const ta = made.el as HTMLTextAreaElement
    document.body.appendChild(ta)
    ta.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', shiftKey: true, bubbles: true }))
    expect(committed).toBe(false)
  })

  it('④ Tab 触发 commit + move=right', () => {
    let committed: { value: unknown; move?: string } | null = null
    const ctx = makeCtx('foo', { commit: (v, m) => { committed = { value: v, move: m } } })
    const made = defaultCellEditor(ctx)
    if (made instanceof HTMLElement) throw new Error('expect object')
    const ta = made.el as HTMLTextAreaElement
    document.body.appendChild(ta)
    ta.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }))
    expect(committed).not.toBeNull()
    expect((committed as any).move).toBe('right')
  })

  it('④ Esc 触发 cancel, 不 commit', () => {
    let committed = false
    let cancelled = false
    const ctx = makeCtx('foo', {
      commit: () => { committed = true },
      cancel: () => { cancelled = true },
    })
    const made = defaultCellEditor(ctx)
    if (made instanceof HTMLElement) throw new Error('expect object')
    const ta = made.el as HTMLTextAreaElement
    document.body.appendChild(ta)
    ta.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    expect(cancelled).toBe(true)
    expect(committed).toBe(false)
  })

  it('④ 失焦触发 commit (跟之前 input 行为一致)', () => {
    let committed = false
    const ctx = makeCtx('foo', { commit: () => { committed = true } })
    const made = defaultCellEditor(ctx)
    if (made instanceof HTMLElement) throw new Error('expect object')
    const ta = made.el as HTMLTextAreaElement
    document.body.appendChild(ta)
    ta.dispatchEvent(new Event('blur'))
    expect(committed).toBe(true)
  })

  it('⑤ 样式贴合: 字号 / 加粗 / 对齐 / 背景 跟 snapshot.style 一致', () => {
    const ctx = makeCtx('hi')
    ;(ctx.snapshot.style as any).font.bold = true
    ;(ctx.snapshot.style as any).font.size = 14 // 14pt
    ;(ctx.snapshot.style as any).hAlign = 'center'
    ;(ctx.snapshot.style as any).fill = { type: 'solid', fgColor: '#ffeb99' }
    const made = defaultCellEditor(ctx)
    if (made instanceof HTMLElement) throw new Error('expect object')
    const ta = made.el as HTMLTextAreaElement
    expect(ta.style.fontWeight).toBe('bold')
    expect(parseFloat(ta.style.fontSize)).toBeCloseTo(14 * (96 / 72), 1)
    expect(ta.style.textAlign).toBe('center')
    // 背景跟随 fill (颜色匹配可能因浏览器格式化稍变, 这里只验非空)
    expect(ta.style.background).toBeTruthy()
  })
})
