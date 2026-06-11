/**
 * 透视表行分组折叠/展开按钮(canvas 绘制 + 控制器命中测试),与 autofilter 下拉按钮同款套路:
 * 画在分组表头行最左格内,点击由 ViewerController.onMouseDown 经 CanvasRenderer.pivotToggleAt 命中。
 * 只在多行字段(外层分组可折叠)时出现,导出时不画(导出件靠真 OOXML 透视表自身的展开)。
 */

/** 折叠按钮方框:左对齐贴在分组表头格内、垂直居中;格太窄(<8px)返回 null 不画。 */
export function pivotToggleBox(cellX: number, cellY: number, cellW: number, cellH: number): { x: number; y: number; size: number } | null {
  const size = Math.min(13, cellW - 2, cellH - 2)
  if (size < 8) return null
  return { x: cellX + 1, y: cellY + (cellH - size) / 2, size }
}

/** 画一个折叠按钮:方框 + 横杠(展开态显示 −)/ 横竖杠(折叠态显示 +)。 */
export function drawPivotToggle(ctx: CanvasRenderingContext2D, x: number, y: number, size: number, collapsed: boolean): void {
  ctx.save()
  ctx.fillStyle = '#ffffff'
  ctx.strokeStyle = '#7a8694'
  ctx.lineWidth = 1
  ctx.fillRect(x + 0.5, y + 0.5, size - 1, size - 1)
  ctx.strokeRect(x + 0.5, y + 0.5, size - 1, size - 1)
  ctx.strokeStyle = '#1f2329'
  ctx.beginPath()
  const cx = x + size / 2
  const cy = y + size / 2
  const arm = size / 2 - 3
  ctx.moveTo(cx - arm, cy)
  ctx.lineTo(cx + arm, cy)
  if (collapsed) {
    ctx.moveTo(cx, cy - arm)
    ctx.lineTo(cx, cy + arm)
  }
  ctx.stroke()
  ctx.restore()
}
