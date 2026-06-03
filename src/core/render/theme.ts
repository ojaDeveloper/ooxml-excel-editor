/** 预览器外观主题(可由 :theme prop 覆盖)。 */
export interface ViewerTheme {
  /** 行/列表头背景 */
  headerBg: string
  /** 选中行/列时表头高亮背景(预留) */
  headerActiveBg: string
  /** 表头文字色 */
  headerText: string
  /** 表头边线色 */
  headerLine: string
  /** 网格线色 */
  gridLine: string
  /** 选区边框色 */
  selBorder: string
  /** 多格选区填充色(含透明) */
  selFill: string
}

export const DEFAULT_THEME: ViewerTheme = {
  headerBg: '#F5F6F7',
  headerActiveBg: '#E1E6EB',
  headerText: '#4B4B4B',
  headerLine: '#C6CCD2',
  gridLine: '#E0E2E5',
  selBorder: '#1A73E8',
  selFill: 'rgba(26,115,232,0.10)',
}

export function mergeTheme(partial?: Partial<ViewerTheme>): ViewerTheme {
  return { ...DEFAULT_THEME, ...(partial || {}) }
}
