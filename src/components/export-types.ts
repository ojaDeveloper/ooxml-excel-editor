/** ExportDialog 表单配置(组件间共享;不能在 <script setup> 里 export 类型) */
export interface ExportConfig {
  action: 'png' | 'pdf' | 'print'
  scope: 'selection' | 'sheet' | 'all'
  scale: number
  includeHeaders: boolean
  gridlines: boolean
  /** 'auto' = 不覆盖,沿用工作表 pageSetup */
  format: 'auto' | 'a4' | 'a3' | 'letter'
  orientation: 'auto' | 'portrait' | 'landscape'
  fitToWidth: boolean
}
