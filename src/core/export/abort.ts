/**
 * 长任务防假死 + 取消的两件小工具(框架无关)。
 *
 * 长导出(PDF / PNG / XLSX、批量图片转换)必须在循环中:
 *   1. checkAborted(signal) — 用户调 abortController.abort() 立刻中断
 *   2. await yieldToEvent() — 让出主线程跑一次绘制 / 事件,避免 UI 假死
 * 配合 ExportProgress 的 onProgress 回调,使用者可显示进度条 + 取消按钮。
 */

/**
 * 让出主线程跑一帧。优先 requestAnimationFrame(浏览器),其次 setTimeout(node / worker 无 rAF)。
 * 用 await yieldToEvent() 包在循环里防 UI 假死。返回 0 让调用方可选取用时间戳。
 */
export function yieldToEvent(): Promise<number> {
  if (typeof requestAnimationFrame === 'function') {
    return new Promise<number>((resolve) => requestAnimationFrame((t) => resolve(t)))
  }
  return new Promise<number>((resolve) => setTimeout(() => resolve(0), 0))
}

/**
 * 抛 AbortError 让上层 try/catch 区分"取消"与"真出错"。
 * AbortSignal 标准:`signal.aborted === true` 时调用方应 throw `DOMException('Aborted', 'AbortError')`。
 * 跨环境兜底:Node 18+ / 现代浏览器 / Worker 均有 DOMException;无 DOMException 时回退普通 Error 但保留 name。
 */
export function checkAborted(signal?: AbortSignal): void {
  if (!signal?.aborted) return
  // DOMException 在 Node 17+ 与浏览器普遍可用;tsc lib.dom 提供类型
  if (typeof DOMException !== 'undefined') {
    throw new DOMException('Aborted', 'AbortError')
  }
  const err = new Error('Aborted')
  ;(err as { name: string }).name = 'AbortError'
  throw err
}

/** 判断是否 AbortError(便于上层在 catch 里区分,避免吞掉真异常) */
export function isAbortError(e: unknown): boolean {
  return !!e && typeof e === 'object' && (e as { name?: string }).name === 'AbortError'
}
