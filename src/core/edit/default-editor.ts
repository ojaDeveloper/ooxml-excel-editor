/**
 * 内置默认单元格编辑器(框架无关 DOM)—— editor 扩展点的默认实现.
 *
 * Phase 1 (2026-06-08, WPS 长文本撑高):
 * - 用 `<textarea>` 代替 `<input>` , 单行场景视觉一致 (initial rows=1), 长文本自动换行 + 撑高
 * - 实现 `getDesiredHeight(width)` , host 据此撑高编辑框向下溢出原格 (跟 WPS 一致)
 * - 提交后**不动行高** (跟 WPS 一致, 不持久化撑高;cell.wrapText=true 时走已有 autofit)
 *
 * 快捷键: Enter 提交+下移, Tab 提交+右移, Esc 取消, Shift+Enter 插入换行, 失焦提交.
 */
import type { CellEditorContext, CellEditorFactory } from './editor-context'
import { CELL_PADDING, LINE_HEIGHT_FACTOR, fontToCss, wrapLines } from '../render/text'
import { FormulaAutocomplete } from './formula-autocomplete'

const CSS_PT_TO_PX = 96 / 72

/** 模块单例: 一份离屏 canvas 给 wrapLines/measureText 用 (textarea 内容变化时高频调用) */
let _measureCtx: CanvasRenderingContext2D | null = null
function getMeasureCtx(): CanvasRenderingContext2D | null {
  if (_measureCtx) return _measureCtx
  if (typeof document === 'undefined') return null
  const c = document.createElement('canvas')
  c.width = 100
  c.height = 100
  _measureCtx = c.getContext('2d')
  return _measureCtx
}

export const defaultCellEditor: CellEditorFactory = (ctx: CellEditorContext) => {
  const st = ctx.snapshot.style
  // 背景跟随单元格填充(像 WPS:编辑时不变白),solid 取 fgColor,pattern 取底色,无填充才白
  const fill = st?.fill
  const bg = fill && fill.type !== 'none' ? (fill.fgColor ?? fill.bgColor ?? '#fff') : '#fff'

  const ta = document.createElement('textarea')
  ta.className = 'ooxml-cell-editor'
  ta.value = ctx.initialText ?? ctx.snapshot.text
  ta.rows = 1
  // resize:none — 用户不能手拖大小, 由 host 控制
  // overflow:auto — 撑到上限后内部滚 (不再向下溢出)
  // whiteSpace:pre-wrap — 保留 \n 同时自动换行 (跟 wrapLines 行为一致)
  // 注: padding 用 0 3px (左右 3 = CELL_PADDING; 上下 0, 行盒高度由 line-height 撑)
  ta.style.cssText =
    'box-sizing:border-box;border:2px solid #21a366;outline:none;padding:0 3px;margin:0;font-family:sans-serif;resize:none;overflow:auto;white-space:pre-wrap;word-break:break-word;'
  ta.style.background = bg

  // 贴合单元格样式(字号/粗斜/对齐/颜色)
  const fontPx = st?.font?.size ? st.font.size * CSS_PT_TO_PX : 14
  ta.style.fontSize = fontPx + 'px'
  ta.style.lineHeight = LINE_HEIGHT_FACTOR.toString() // 跟渲染层一致, 撑高算出来才能跟 textarea 实际高度对齐
  if (st?.font?.bold) ta.style.fontWeight = 'bold'
  if (st?.font?.italic) ta.style.fontStyle = 'italic'
  if (st?.font?.color) ta.style.color = st.font.color
  ta.style.textAlign = st?.hAlign === 'center' ? 'center' : st?.hAlign === 'right' ? 'right' : 'left'

  // 公式自动补全(1.14.0):输 =SU 弹函数列表。列表打开时拦 ↑↓/Enter/Tab/Esc。
  const ac = new FormulaAutocomplete(ta)
  let done = false
  // 上一次被拒的值:blur(点提示弹窗)会再触发 commit,同一被拒值不重复提交 → 避免弹窗叠弹
  let lastRejected: string | null = null
  const commit = (move?: 'down' | 'right') => {
    if (done) return
    if (ta.value === lastRejected) return // 同一被拒值不重复提交(Enter/blur 二次触发)
    done = true
    ac.dispose()
    // commit 返 false = 被拒(数据验证拦截等)→ 解除锁,记住被拒值,编辑器留开让用户改正
    if (ctx.commit(ta.value, move) === false) { done = false; lastRejected = ta.value }
  }
  const cancel = () => {
    if (done) return
    done = true
    ac.dispose()
    ctx.cancel()
  }

  ta.addEventListener('keydown', (e) => {
    if (ac.onKeyDown(e)) { e.preventDefault(); e.stopPropagation(); return } // 补全列表优先吃导航键
    if (e.key === 'Enter' && !e.shiftKey) {
      // 普通 Enter 提交(跟 input 行为一致). Shift+Enter 插入换行 (textarea 原生)
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
  ta.addEventListener('blur', () => commit())

  // 内容变化 → 清除"被拒值"记忆(改了内容就能再次提交) + 公式补全重算 + 通知 host 重撑高 (Phase 1 撑高核心入口)
  ta.addEventListener('input', () => {
    lastRejected = null
    ac.update()
    ctx.reposition?.()
  })

  /**
   * 算"按列宽撑开 N 行需要多高". 跟 canvas 渲染层用同一份 wrapLines (text.ts),
   * 行高 = fontPx * LINE_HEIGHT_FACTOR, 加上下 padding (textarea 的 line-height 已设为
   * LINE_HEIGHT_FACTOR, 所以 N 行高度 = N * fontPx * LINE_HEIGHT_FACTOR).
   */
  function computeDesiredHeight(widthPx: number): number {
    const measureCtx = getMeasureCtx()
    if (!measureCtx || widthPx <= 0) return 0
    const fontCss = st?.font ? fontToCss(st.font, 1) : `${fontPx.toFixed(1)}px sans-serif`
    // 减去左右 padding (3+3) 和边框 (2+2) = 10
    const innerWidth = Math.max(10, widthPx - 2 * CELL_PADDING - 4)
    // wrapLines 接的是当前 textarea 文本; 空文本时给个空格保证至少 1 行高度
    const lines = wrapLines(measureCtx, ta.value || ' ', fontCss, innerWidth)
    const lineH = fontPx * LINE_HEIGHT_FACTOR
    // 高度 = 行数 * 行盒高 + 上下边框 (2+2=4) + 微小余量防最后一行被裁
    return Math.ceil(lines.length * lineH + 4 + 2)
  }

  // 打字进入时光标置末尾;F2/双击进入时全选(便于整体替换)
  setTimeout(() => {
    ta.focus()
    if (ctx.initialText != null) ta.setSelectionRange(ta.value.length, ta.value.length)
    else ta.select()
    // 焦点 + 初始内容到位后, 主动撑一次 (长文本进入编辑时立即撑高)
    ctx.reposition?.()
  }, 0)

  return { el: ta, getDesiredHeight: computeDesiredHeight }
}
