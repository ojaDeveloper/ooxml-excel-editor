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
  // 透视表: 小表格 + 汇总箭头
  'pivot-table': '<rect x="3" y="4" width="13" height="13" rx="1.5"/><line x1="3" y1="8" x2="16" y2="8"/><line x1="7" y1="4" x2="7" y2="17"/><path d="M14 14h7"/><path d="M18 11l3 3-3 3"/>',
  // 条件格式: 渐变色块 + 小滴
  'conditional-format': '<rect x="3" y="4" width="13" height="13" rx="1.5"/><path d="M3 9h13"/><path d="M3 13h13"/><circle cx="18.5" cy="15" r="3"/><path d="M18.5 9.5c1.2 1.4 2 2.7 2 3.8a2 2 0 0 1-4 0c0-1.1.8-2.4 2-3.8z"/>',
  // 数字格式: 百分号 + 小数点(123)
  'number-format': '<text x="2" y="11" font-size="9" font-family="sans-serif" fill="currentColor" stroke="none">123</text><circle cx="5" cy="17" r="1" fill="currentColor" stroke="none"/><path d="M21 6l-7 12"/><circle cx="14.5" cy="8" r="2"/><circle cx="20.5" cy="16" r="2"/>',
  // 下载(导出)
  export: '<line x1="12" y1="3" x2="12" y2="15"/><path d="M8 11l4 4 4-4"/><path d="M4 19h16"/>',
  // 放大镜带 +
  zoom: '<circle cx="10" cy="10" r="6"/><line x1="20" y1="20" x2="14.5" y2="14.5"/><line x1="10" y1="7.5" x2="10" y2="12.5"/><line x1="7.5" y1="10" x2="12.5" y2="10"/>',
  // 两个叠框(复制)
  copy: '<rect x="9" y="9" width="11" height="11" rx="1.5"/><path d="M5 15H4V5a1 1 0 0 1 1-1h10v1"/>',
  // 十字线(冻结行列)
  freeze: '<line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="3" x2="9" y2="21"/>',
  // 自动换行:两条文本线 + 折回箭头
  'wrap-text':
    '<line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="18" x2="10" y2="18"/>' +
    '<path d="M3 12h15a3 3 0 0 1 0 6h-3"/><polyline points="17 15 14 18 17 21"/>',
  // 模板:文档 + 排版网格
  template:
    '<rect x="4" y="3" width="16" height="18" rx="1.5"/>' +
    '<line x1="4" y1="8" x2="20" y2="8"/>' +
    '<line x1="8" y1="12" x2="16" y2="12"/>' +
    '<line x1="8" y1="16" x2="16" y2="16"/>',
  // 图片工具:山+太阳的图片框 + 小齿轮(批量转换)
  'image-tools':
    '<rect x="3" y="4" width="14" height="14" rx="1.5"/><circle cx="8" cy="9" r="1.5"/>' +
    '<polyline points="3 16 7 12 11 16 14 13 17 16"/>' +
    '<path d="M19 19l2 2M19 21l2-2"/>',
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
