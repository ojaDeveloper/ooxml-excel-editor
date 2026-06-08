/**
 * 中间数据模型 —— 渲染层只认这套类型，与 ExcelJS / 原始 XML 形状完全解耦。
 * 将来要换解析库或换 DOM 渲染，只要保证产出/消费这套模型即可。
 */

/** RGB(A) 颜色，统一成 css 可用字符串(如 "#RRGGBB" / "rgba(...)")。 */
export type CssColor = string

/** 水平对齐 */
export type HAlign = 'left' | 'center' | 'right' | 'fill' | 'justify' | 'general'
/** 垂直对齐 */
export type VAlign = 'top' | 'middle' | 'bottom'

/** 单个边框边的样式 */
export interface BorderEdge {
  style: BorderStyle
  color: CssColor
}
export type BorderStyle =
  | 'none'
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot'

export interface Borders {
  top?: BorderEdge
  bottom?: BorderEdge
  left?: BorderEdge
  right?: BorderEdge
  /** 对角线边框样式(up/down 共用此线型与颜色) */
  diagonal?: BorderEdge
  /** ↗ 左下→右上 对角线 */
  diagonalUp?: boolean
  /** ↘ 左上→右下 对角线 */
  diagonalDown?: boolean
}

export interface Font {
  name: string
  size: number // pt
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  color: CssColor
}

export type FillType = 'none' | 'solid' | 'pattern' | 'gradient'

export interface Fill {
  type: FillType
  /** solid / pattern 前景色；gradient 时为第一个 stop */
  fgColor?: CssColor
  bgColor?: CssColor
  pattern?: string // e.g. "gray125"
  /** 渐变色标(可选，简化版) */
  gradientStops?: { position: number; color: CssColor }[]
  gradientAngle?: number
}

export interface CellStyle {
  font: Font
  fill: Fill
  borders: Borders
  hAlign: HAlign
  vAlign: VAlign
  wrapText: boolean
  shrinkToFit: boolean
  /** 文本旋转角度(度)，Excel 范围 -90..90，255 表示竖排 */
  textRotation: number
  indent: number
  /** 数字格式代码字符串(已从 numFmtId 解析为 code) */
  numFmt: string
}

/** 单元格的逻辑值类型 */
export type CellValueType =
  | 'empty'
  | 'number'
  | 'string'
  | 'boolean'
  | 'date'
  | 'error'
  | 'hyperlink'
  | 'richtext'
  | 'formula'

export interface RichTextRun {
  text: string
  font?: Partial<Font>
}

export interface CellModel {
  row: number // 0-based
  col: number // 0-based
  type: CellValueType
  /** 原始值: number / string / boolean / Date / error 字符串 */
  raw: number | string | boolean | Date | null
  /** 富文本分段(type === 'richtext' 时有) */
  rich?: RichTextRun[]
  /** 公式文本(type === 'formula' 时有；显示用 raw 缓存值) */
  formula?: string
  /** 超链接目标 */
  hyperlink?: string
  /** 单元格批注(纯文本) */
  comment?: string
  /** 样式索引，指向 SheetModel.styles */
  styleId: number
  /**
   * WPS DISPIMG 单元格内嵌图的 id(指向 WorkbookModel.cellImages)。
   * 有值时渲染器把图画进格内(随行高列宽缩放、随网格滚动裁剪),而非画 formula 文本。
   * 由 cell-image-parser 从 `=DISPIMG("id",n)` 公式解析填充。
   */
  dispImgId?: string
}

/** 合并区域(0-based，闭区间) */
export interface MergeRange {
  top: number
  left: number
  bottom: number
  right: number
}

export interface ColumnInfo {
  /** 列宽，单位 = px(已从字符数换算) */
  width: number
  hidden: boolean
}
export interface RowInfo {
  /** 行高，单位 = px(已从 pt 换算) */
  height: number
  hidden: boolean
  /**
   * 是否"手动设定行高"(OOXML `<row customHeight="1">`)。
   * true 时渲染层**不做自动行高**(与 Excel/WPS 一致:手动高度的行只裁切/溢出,不撑大)。
   * 由 row-meta-parser 从原始 XML 读出(ExcelJS 不暴露此标记)。
   */
  customHeight?: boolean
}

/** 冻结窗格: 冻结前 frozenRows 行 / frozenCols 列 */
export interface FreezeInfo {
  frozenRows: number
  frozenCols: number
}

/** 条件格式规则(简化版，覆盖常见 4 类) */
export interface ConditionalRule {
  ranges: MergeRange[]
  priority: number
  type: 'cellIs' | 'colorScale' | 'dataBar' | 'iconSet' | 'expression' | 'top10' | 'unsupported'
  /** cellIs */
  operator?: string
  formulae?: string[]
  /** 命中时套用的样式(cellIs / expression) */
  style?: Partial<CellStyle>
  /** colorScale: 2~3 个色标 */
  colorScale?: { min: CssColor; mid?: CssColor; max: CssColor }
  /** dataBar */
  dataBar?: { color: CssColor; gradient: boolean }
  /** iconSet */
  iconSet?: { name: string }
}

/** 图片锚定(像素矩形由 layout 阶段最终算出，这里给逻辑锚点) */
export interface ImageAnchor {
  /** blob url(由主线程从 bytes 生成；解析阶段为空) */
  src: string
  /** 原始图片字节 + mime(解析层产出，可跨 Worker 传输) */
  bytes?: Uint8Array
  mime?: string
  /** twoCellAnchor: from/to 单元格 + 单元格内偏移(EMU) */
  from: AnchorCell
  to?: AnchorCell
  /** oneCellAnchor / absolute 的尺寸(EMU) */
  extWidthEmu?: number
  extHeightEmu?: number
  editAs?: string
}
export interface AnchorCell {
  col: number // 0-based
  colOffEmu: number
  row: number // 0-based
  rowOffEmu: number
}

/**
 * WPS DISPIMG 单元格内嵌图(workbook 级登记表 xl/cellimages.xml 的一条)。
 * 单元格通过 `=DISPIMG("id",n)` 公式按 id 引用;不同于浮动图(ImageAnchor),它"属于"单元格。
 */
export interface CellImage {
  /** DISPIMG id(= cellimages.xml 里 cNvPr@name,如 "ID_5db4b3...") */
  id: string
  /** 原始图片字节 + mime(解析层产出,可跨 Worker 传输) */
  bytes?: Uint8Array
  mime?: string
  /** blob url(主线程 finalize 从 bytes 生成;解析阶段为空) */
  src: string
}

/** 图表规格(从 chartN.xml 抽出，交给 ECharts 映射) */
export interface ChartSpec {
  type: 'bar' | 'line' | 'pie' | 'area' | 'scatter' | 'doughnut' | 'radar' | 'unsupported'
  title?: string
  showLegend: boolean
  barDirection?: 'col' | 'bar' // 柱状: 垂直 / 水平
  categories: (string | number)[]
  series: ChartSeries[]
  anchor: ImageAnchor // 复用锚定结构定位
}
export interface ChartSeries {
  name?: string
  values: (number | null)[]
  color?: CssColor
}

export interface SheetModel {
  name: string
  index: number
  state: 'visible' | 'hidden' | 'veryHidden'
  /** 实际数据范围 */
  dimension: { rows: number; cols: number }
  /** 稀疏存储: key = `${row}:${col}` */
  cells: Map<string, CellModel>
  /** 样式表，CellModel.styleId 索引 */
  styles: CellStyle[]
  merges: MergeRange[]
  columns: Map<number, ColumnInfo> // 0-based col -> info
  rows: Map<number, RowInfo> // 0-based row -> info
  defaultColWidth: number // px
  defaultRowHeight: number // px
  freeze: FreezeInfo
  conditional: ConditionalRule[]
  /** 自动筛选区域(画下拉按钮用) */
  autoFilterRange?: MergeRange
  /** 含"列表"型数据验证的区域(选中时画下拉箭头) */
  dataValidations: MergeRange[]
  images: ImageAnchor[]
  charts: ChartSpec[]
  /** 形状 / 文本框(DrawingML sp) */
  shapes: ShapeSpec[]
  /** 迷你图(单元格内嵌折线/柱/盈亏图) */
  sparklines: Sparkline[]
  /** 手动分页符(0-based 边界索引): 在这些行上方/列左侧画分页虚线 */
  pageBreaks?: { rows: number[]; cols: number[] }
  /** 原生页面设置(打印/导出默认值来源);缺省走 export 模块默认 */
  pageSetup?: PageSetupModel
  showGridLines: boolean
}

/** OOXML 原生页面设置(pageSetup + pageMargins + 打印区域/标题),用作导出默认值 */
export interface PageSetupModel {
  orientation?: 'portrait' | 'landscape'
  /** 纸张: 由 paperSize 代码映射;无法识别时省略(走默认 a4) */
  paperFormat?: 'a4' | 'a3' | 'letter' | [number, number]
  /** 缩放百分比(如 80 = 80%);fitToPage 时无意义 */
  scale?: number
  /** 适应页面(fitToPage): 优先于 scale */
  fitToPage?: boolean
  fitToWidth?: number
  fitToHeight?: number
  /** 页边距 mm(从 inch 换算) */
  margins?: { top: number; right: number; bottom: number; left: number; header: number; footer: number }
  /** 打印区域(0-based 闭区间;多区域取第一段) */
  printArea?: MergeRange
  /** 打印标题行 [r0,r1] 0-based(每页顶部重复) */
  printTitleRows?: [number, number]
  /** 打印标题列 [c0,c1] 0-based(横向分页时左侧重复;当前竖向分页管线不应用) */
  printTitleCols?: [number, number]
}

/** 形状 / 文本框: 用锚点定位,带填充/边框/文字 */
export interface ShapeSpec {
  anchor: ImageAnchor
  /** 形状类型(prstGeom): rect / roundRect / ellipse / 其它(按 rect 处理) */
  geom: 'rect' | 'roundRect' | 'ellipse' | 'other'
  text?: string
  fillColor?: CssColor
  lineColor?: CssColor
  textColor?: CssColor
  bold?: boolean
  /** 文本水平对齐 */
  align?: 'left' | 'center' | 'right'
}

/** 迷你图: 锚在某个单元格，数据来自一段区域的数值 */
export interface Sparkline {
  row: number // 0-based 位置
  col: number
  type: 'line' | 'column' | 'winloss'
  values: (number | null)[]
  /** 主色(可选，缺省用默认蓝) */
  color?: CssColor
  negativeColor?: CssColor
}

export interface WorkbookModel {
  sheets: SheetModel[]
  activeSheet: number
  /** 主题色调色板(theme1.xml 解析所得，索引同 ECMA dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink) */
  themeColors: CssColor[]
  date1904: boolean
  /**
   * WPS DISPIMG 单元格内嵌图登记表(id → 图)。对应 xl/cellimages.xml,workbook 级共享。
   * 单元格 CellModel.dispImgId 指向此表;无 WPS 内嵌图时为 undefined。
   */
  cellImages?: Map<string, CellImage>
}

export const cellKey = (row: number, col: number) => `${row}:${col}`

/** 数据钩子: 解析后、渲染前改模型(返回新模型或就地改) */
export type TransformModelFn = (workbook: WorkbookModel) => WorkbookModel | void

/** 单元格样式覆盖(各字段可选;font/fill/borders 允许部分,与解析样式浅合并) */
export interface CellStyleOverride {
  font?: Partial<Font>
  fill?: Partial<Fill>
  borders?: Partial<Borders>
  hAlign?: HAlign
  vAlign?: VAlign
  wrapText?: boolean
  shrinkToFit?: boolean
  textRotation?: number
  indent?: number
  numFmt?: string
}
/**
 * 渲染钩子上下文 (Phase C, 2026-06-08).
 * 给 `cellStyle` 钩子的第 3 入参, 让插件能感知该格当前是否可编辑 →
 * 给只读格定制样式不再需要在 setup 里间接调 viewer.isCellEditable.
 */
export interface CellStyleCtx {
  /** 该格此刻是否可编辑(综合 editable + editableTargets + readOnlyRanges + cellReadOnly) */
  editable: boolean
}
/**
 * 渲染钩子: 按单元格覆盖样式(返回部分样式,与解析样式合并).
 * Phase C 2026-06-08: 加可选第 3 入参 `ctx: CellStyleCtx`, 含 `editable`. 旧 `(cell, pos) => ...` 签名兼容.
 */
export type CellStyleFn = (cell: CellModel, pos: { row: number; col: number }, ctx?: CellStyleCtx) => CellStyleOverride | void
