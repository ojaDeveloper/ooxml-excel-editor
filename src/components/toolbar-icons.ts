/**
 * 极简线性图标(24×24 viewBox,stroke=currentColor)。内置工具栏项按名引用,跨平台一致。
 * 值为 <svg> 内部标记;渲染见 ActionToolbar。
 */
export const TOOLBAR_ICONS: Record<string, string> = {
  // 放大镜
  find: '<circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>',
  // 漏斗
  filter: '<path d="M3 5h18l-7 8.5V20l-4-2.2V13.5z"/>',
  // 漏斗 + 斜杠(清除筛选)
  'clear-filter': '<path d="M3 5h18l-7 8.5V20l-4-2.2V13.5z"/><line x1="3" y1="3" x2="21" y2="21"/>',
  // 上下箭头(排序)
  sort: '<path d="M7 16l3 3 3-3"/><line x1="10" y1="19" x2="10" y2="5"/><path d="M17 8l-3-3-3 3"/><line x1="14" y1="5" x2="14" y2="19"/>',
  // 下载(导出)
  export: '<line x1="12" y1="3" x2="12" y2="15"/><path d="M8 11l4 4 4-4"/><path d="M4 19h16"/>',
  // 放大镜带 +
  zoom: '<circle cx="10" cy="10" r="6"/><line x1="20" y1="20" x2="14.5" y2="14.5"/><line x1="10" y1="7.5" x2="10" y2="12.5"/><line x1="7.5" y1="10" x2="12.5" y2="10"/>',
  // 两个叠框(复制)
  copy: '<rect x="9" y="9" width="11" height="11" rx="1.5"/><path d="M5 15H4V5a1 1 0 0 1 1-1h10v1"/>',
  // 十字线(冻结行列)
  freeze: '<line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="3" x2="9" y2="21"/>',
  // 三点(更多)
  more: '<circle cx="5" cy="12" r="1.6" fill="currentColor" stroke="none"/><circle cx="12" cy="12" r="1.6" fill="currentColor" stroke="none"/><circle cx="19" cy="12" r="1.6" fill="currentColor" stroke="none"/>',
  // 下拉小箭头
  caret: '<path d="M6 9l6 6 6-6"/>',
}

/** 把图标内部标记包成完整 <svg>(16×16,stroke=currentColor) */
export function svgWrap(inner: string): string {
  return (
    '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" ' +
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
    inner +
    '</svg>'
  )
}
