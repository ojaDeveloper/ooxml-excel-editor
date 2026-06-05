/**
 * demo 用的自定义编辑器(框架无关 DOM)—— Vue / React demo 共用,演示 E2 的 editor 扩展点。
 * 第 0 列(产品)用一个 <select> 下拉编辑;其余列无自定义编辑器(E3 内置文本编辑器接管)。
 */
import type { EditorResolver } from '@/core/edit/editor-context'

export const demoSelectEditor: EditorResolver = (_cell, pos) => {
  if (pos.col !== 0) return
  return (ctx) => {
    const sel = document.createElement('select')
    sel.className = 'demo-cell-editor'
    const cur = ctx.snapshot.text
    for (const v of [cur, 'AAA', 'BBB'].filter((x, i, a) => a.indexOf(x) === i)) {
      const o = document.createElement('option')
      o.value = v
      o.textContent = v || '(空)'
      sel.appendChild(o)
    }
    sel.value = cur
    sel.style.cssText = 'font:13px sans-serif;border:1px solid #21a366;outline:none;'
    sel.onchange = () => ctx.commit(sel.value)
    sel.onkeydown = (e) => {
      if (e.key === 'Escape') ctx.cancel()
    }
    setTimeout(() => sel.focus(), 0)
    return sel
  }
}
