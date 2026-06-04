/** 操作栏解析后的按钮(ExcelViewer 算好 handler/active,ActionToolbar 只管渲染) */
export interface ResolvedToolbarItem {
  id: string
  icon?: string
  label?: string
  title?: string
  active?: boolean
  onClick: () => void
  /** 来源: 内置 / 组件配置自定义 / 插件 —— 用于分组分隔线 */
  kind: 'builtin' | 'custom' | 'plugin'
}
