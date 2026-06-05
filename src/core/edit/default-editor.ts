/**
 * 内置默认单元格编辑器(框架无关 DOM)—— editor 扩展点的默认实现:一个贴合单元格样式的 <input>。
 * Enter 提交+下移,Tab 提交+右移,Esc 取消,失焦提交。用方/插件给了 editor 就用它们的,否则用这个。
 */
import type { CellEditorContext, CellEditorFactory } from './editor-context'

const CSS_PT_TO_PX = 96 / 72

export const defaultCellEditor: CellEditorFactory = (ctx: CellEditorContext) => {
  const st = ctx.snapshot.style
  // 背景跟随单元格填充(像 WPS:编辑时不变白),solid 取 fgColor,pattern 取底色,无填充才白
  const fill = st?.fill
  const bg = fill && fill.type !== 'none' ? (fill.fgColor ?? fill.bgColor ?? '#fff') : '#fff'

  const input = document.createElement('input')
  input.type = 'text'
  input.className = 'ooxml-cell-editor'
  input.value = ctx.initialText ?? ctx.snapshot.text
  input.style.cssText =
    'box-sizing:border-box;border:2px solid #21a366;outline:none;padding:0 3px;margin:0;font-family:sans-serif;'
  input.style.background = bg

  // 贴合单元格样式(字号/粗斜/对齐/颜色)
  input.style.fontSize = (st?.font?.size ? st.font.size * CSS_PT_TO_PX : 14) + 'px'
  if (st?.font?.bold) input.style.fontWeight = 'bold'
  if (st?.font?.italic) input.style.fontStyle = 'italic'
  if (st?.font?.color) input.style.color = st.font.color
  input.style.textAlign = st?.hAlign === 'center' ? 'center' : st?.hAlign === 'right' ? 'right' : 'left'

  let done = false
  const commit = (move?: 'down' | 'right') => {
    if (done) return
    done = true
    ctx.commit(input.value, move)
  }
  const cancel = () => {
    if (done) return
    done = true
    ctx.cancel()
  }

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault()
      commit('down')
    } else if (e.key === 'Tab') {
      e.preventDefault()
      commit('right')
    } else if (e.key === 'Escape') {
      e.preventDefault()
      cancel()
    }
    e.stopPropagation() // 别让网格的键盘处理插手
  })
  input.addEventListener('blur', () => commit())

  // 打字进入时光标置末尾;F2/双击进入时全选(便于整体替换)
  setTimeout(() => {
    input.focus()
    if (ctx.initialText != null) input.setSelectionRange(input.value.length, input.value.length)
    else input.select()
  }, 0)

  return { el: input }
}
