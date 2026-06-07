/**
 * 导出/打印的公共选项类型。
 * 设计为"可扩展": PDF 暴露 beforeRenderPage 命令式钩子(页眉/页脚/水印/页码),
 * 打印暴露 headerHtml/footerHtml。宿主与插件都能借此二次定制而不改源码。
 */
import type { MergeRange } from '../model/types'
import type { ExportProgressFn } from '../progress'

/** 选哪些工作表导出 */
export type ExportTarget = 'active' | 'all' | number | number[]

/** 公共渲染选项(决定底图怎么画) */
export interface RenderExportOptions {
  /** 仅当导出单个工作表时生效: 限定单元格区域(0-based 闭区间) */
  range?: MergeRange
  /** 设备像素缩放,越大越清晰(默认 2) */
  scale?: number
  /** 含行号/列字母表头(默认 false,同 Excel 打印) */
  includeHeaders?: boolean
  /** 覆盖网格线显隐(缺省跟随工作表设置) */
  gridlines?: boolean
  /** 背景色(默认白) */
  background?: string
  /**
   * 长任务进度回调。每个调度阶段(render/compose/paginate/write/zip)各 emit。
   * 大表 / 多表 / 多页时穿插 `await yieldToEvent()` 避免 UI 假死。
   */
  onProgress?: ExportProgressFn
  /** 取消信号。`abortController.abort()` 后下一个调度点抛 AbortError(标准语义) */
  signal?: AbortSignal
}

/** 图片导出选项 */
export interface ImageExportOptions extends RenderExportOptions {
  target?: ExportTarget // 默认 'active'(仅单表能直接出一张图)
  /** 图片格式(默认 png) */
  type?: 'png' | 'jpeg' | 'webp'
  /** jpeg/webp 质量 0~1(默认 0.92) */
  quality?: number
  /** 下载文件名(省略则用 "<表名>.png") */
  fileName?: string
}

// ---------------- 分页 / 纸张 ----------------
export type PageFormat = 'a4' | 'a3' | 'letter' | [number, number] // [宽, 高] mm
export type Orientation = 'portrait' | 'landscape'
export interface Margins {
  top: number
  right: number
  bottom: number
  left: number
} // mm

export interface PageSetup {
  /** 纸张(默认 a4) */
  format?: PageFormat
  /** 方向(默认 portrait) */
  orientation?: Orientation
  /** 页边距 mm(数字=四边相同;默认 10) */
  margin?: number | Partial<Margins>
  /** 把内容缩放到页宽(默认 true);false 则按自然尺寸,超宽会横向分页 */
  fitToWidth?: boolean
}

/** PDF 单页绘制上下文 —— 传给 beforeRenderPage 钩子(在页内容已贴图后调用) */
export interface PdfPageContext {
  /** jsPDF 实例(可调 doc.text / doc.setFontSize / doc.line 等画页眉页脚水印) */
  doc: any
  /** 全局页序(0-based) */
  pageIndex: number
  /** 全局总页数 */
  pageCount: number
  /** 本页所属工作表名 */
  sheetName: string
  /** 本页所属工作表在导出集合中的序(0-based) */
  sheetIndex: number
  /** 整页尺寸 mm */
  pageWidth: number
  pageHeight: number
  /** 页边距 mm */
  margin: Margins
}
export type BeforeRenderPage = (ctx: PdfPageContext) => void

/** PDF 导出选项 */
export interface PdfExportOptions extends RenderExportOptions, PageSetup {
  target?: ExportTarget // 默认 'active'
  /** 下载文件名(省略则用 "<文件名|workbook>.pdf") */
  fileName?: string
  /** 每页贴图后调用,用于画页眉/页脚/水印/页码等(扩展点) */
  beforeRenderPage?: BeforeRenderPage
  /**
   * 矢量 PDF: 逐格用 jsPDF 真文字/矢量绘制(可选可搜、清晰、文件小)。
   * 默认 false = 位图 PDF(整表贴图)。
   */
  vector?: boolean
  /**
   * 矢量模式扩展钩子: 建好 jsPDF doc 后调用一次,宿主可在此 addFont 注册中文等字体。
   *   configureDoc: (doc) => { doc.addFileToVFS('f.ttf', b64); doc.addFont('f.ttf','CN','normal'); doc.setFont('CN') }
   * 注册了自定义字体 → 全矢量;否则非拉丁文本(中文)的单元格自动栅格兜底(不丢内容)。
   */
  configureDoc?: (doc: any) => void
}

/** 打印选项 */
export interface PrintOptions extends RenderExportOptions, PageSetup {
  target?: ExportTarget // 默认 'active'
  /** 打印窗口标题 */
  title?: string
  /** 每页顶部 HTML 片段(如公司抬头);出现在每个打印页 */
  headerHtml?: string
  /** 每页底部 HTML 片段 */
  footerHtml?: string
}
