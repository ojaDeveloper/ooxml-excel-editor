/** 操作栏解析后的按钮(ExcelViewer 算好 handler/active/disabled,ActionToolbar 只管渲染) */
export interface ResolvedToolbarItem {
  id: string
  /** 'button'(默认) | 'separator' */
  type?: 'button' | 'separator'
  /** 内联 SVG 内部标记(优先);否则用 icon 文本/emoji */
  iconSvg?: string
  icon?: string
  label?: string
  title?: string
  active?: boolean
  disabled?: boolean
  onClick?: () => void
  /** 下拉子菜单(已解析) */
  items?: ResolvedToolbarItem[]
  /** 来源: 内置 / 组件配置自定义 / 插件 —— 用于分组分隔线 */
  kind: 'builtin' | 'custom' | 'plugin'
}
