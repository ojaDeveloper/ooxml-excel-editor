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
  /** 稳定 id(解析:`cf-p<n>`;用户新建:`cf-u<n>`)。编辑/删除/导出对账用。1.9.0 起;老数据缺省 */
  id?: string
  /** 来源:'parsed' 从文件解析;'user' app 内新建。overlay 导出据此决定原样回写还是按模型写。缺省按 parsed */
  origin?: 'parsed' | 'user'
  /** app 内被编辑过(parsed 规则改过后置 true)。导出:parsed && !dirty → 原样回写 raw;否则按模型写 */
  dirty?: boolean
  ranges: MergeRange[]
  priority: number
  type: 'cellIs' | 'colorScale' | 'dataBar' | 'iconSet' | 'expression' | 'top10' | 'unsupported'
  /** cellIs */
  operator?: string
  formulae?: string[]
  /** 命中时套用的样式(cellIs / expression / top10)。dxf 各字段都可缺,故 font 也是 Partial */
  style?: Omit<Partial<CellStyle>, 'font'> & { font?: Partial<Font> }
  /** colorScale: 2~3 个色标 */
  colorScale?: { min: CssColor; mid?: CssColor; max: CssColor }
  /** dataBar */
  dataBar?: { color: CssColor; gradient: boolean }
  /** iconSet */
  iconSet?: { name: string; reverse?: boolean }
  /** top10: rank 个 / percent 百分比 / bottom 底部 */
  top10?: { rank: number; percent: boolean; bottom: boolean }
  /**
   * 导出专用:解析时原始 ExcelJS rule 对象(含 cfvo 阈值等我们不全建模的字段)。
   * overlay 导出对"未编辑的 parsed 规则"原样回写保真;编辑色阶/数据条/图标集时尽量改这里的颜色/名称、留住阈值。
   * 框架无关模型刻意只放它作不透明透传载体,渲染/编辑逻辑不读它。
   */
  raw?: unknown
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

/**
 * 数据验证规则(完整版,1.8.0):承载校验语义 —— 编辑时拦截非法输入 + 输入/出错提示。
 * 渲染层只用 list 型画下拉箭头(见 SheetModel.dataValidations / dataValidationLists,从这里派生)。
 * formulae 已尽量解析成字面量:整数/小数/文本长度 → number;日期/时间 → number(序列值)或原始串;
 * list → 选项数组在 options;custom → 公式串(暂不求值,放行)。
 */
export interface DataValidationRule {
  range: MergeRange
  type: 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom'
  /** 比较运算符(whole/decimal/date/time/textLength 用;list/custom 不用) */
  operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual'
  /** 约束操作数(between 用前两个;单目用第一个)。原样保留,校验时按 type 解析 */
  formulae: (string | number)[]
  /** 允许留空(空值不校验) */
  allowBlank: boolean
  /** list 型可选值(= dataValidationLists 同源) */
  options?: string[]
  /** 出错提示(showErrorMessage 时,非法输入弹) */
  showErrorMessage?: boolean
  errorStyle?: 'stop' | 'warning' | 'information'
  errorTitle?: string
  error?: string
  /** 输入提示(showInputMessage 时,选中该格弹气泡) */
  showInputMessage?: boolean
  promptTitle?: string
  prompt?: string
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
  /** 列表型数据验证的可选值(点下拉箭头弹选;range 内任一格命中即用 options)。可选 —— 老数据/无选项时缺省 */
  dataValidationLists?: { range: MergeRange; options: string[] }[]
  /** 完整数据验证规则(校验语义:编辑拦截 + 输入/出错提示)。1.8.0 起;上面两个字段从这里派生 */
  dataValidationRules?: DataValidationRule[]
  images: ImageAnchor[]
  charts: ChartSpec[]
  /** 形状 / 文本框(DrawingML sp) */
  shapes: ShapeSpec[]
  /** 迷你图(单元格内嵌折线/柱/盈亏图) */
  sparklines: Sparkline[]
  /** 透视表只读 UI 元数据:用于叠加字段按钮/下拉箭头;数据仍按普通单元格显示 */
  pivotTables: PivotTableModel[]
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

/** 透视表字段按钮(只读 UI):按钮锚在对应标题/筛选单元格上。 */
export interface PivotButton {
  row: number
  col: number
  label: string
  kind: 'row' | 'col' | 'page' | 'data' | 'field'
}

/** 透视表模型(只读):范围来自 pivotTableDefinition/location,字段来自 cacheFields + axis/dataFields。 */
export interface PivotTableModel {
  name: string
  range: MergeRange
  fields: string[]
  buttons: PivotButton[]
  /** 静态透视表来源:当前模型内的源表 index + 源数据区域。用于运行时重建/后续导出。 */
  source?: { sheetIndex: number; range: MergeRange }
  /** 运行时透视布局元数据:字段 index 均为源数据区域内的绝对列 index。 */
  layout?: PivotTableLayout
  /** 已折叠的外层行分组 key(行字段 ≥2 时,外层分组可折叠隐藏明细;空/缺省 = 全展开)。 */
  collapsed?: string[]
  /** 运行时:可折叠的分组表头所在输出行(绝对行号)+ 外层 key;每次重算刷新,供渲染折叠按钮 + 命中测试。 */
  rowGroups?: { row: number; key: string }[]
}

export type PivotSummary = 'sum' | 'count' | 'avg' | 'max' | 'min'
/** all=全部 / non-empty=非空 / equals=单值等于 / include=多选包含(values 列出保留值,空=不约束)。 */
export type PivotFilterMode = 'all' | 'non-empty' | 'equals' | 'include'
export interface PivotFilterRule { field: number; mode: PivotFilterMode; value?: string; values?: string[] }
export interface PivotValueRule { field: number; summary: PivotSummary }
export interface PivotTableLayout {
  filters: PivotFilterRule[]
  columns: number[]
  rows: number[]
  values: PivotValueRule[]
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

/**
 * 规范的"空白默认样式"——所有 SheetModel.styles[0] 必须是它(中性、无填充、无边框)。
 * 空格 / 新建格 / setCellValue 落的格 / applyStyleOverride 的兜底基样式都回落到 styleId 0,
 * 因此 index 0 绝不能是"恰好第一个被解析到的单元格样式"(否则那个格的底色/边框会冒到所有默认格,
 * 见 parser 把首格 A1 绿底当默认导致粘贴/编辑串色的 bug)。loader-json / parser / clipboard-snapshot 共用此工厂。
 */
export function makeDefaultStyle(): CellStyle {
  return {
    font: { name: 'Calibri', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#000000' },
    fill: { type: 'none' },
    borders: {},
    hAlign: 'general',
    vAlign: 'bottom',
    wrapText: false,
    shrinkToFit: false,
    textRotation: 0,
    indent: 0,
    numFmt: 'General',
  }
}

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
