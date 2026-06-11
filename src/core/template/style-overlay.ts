/**
 * 模板样式 overlay(P3 进阶,重设计 2026-06-08)
 * ────────────────────────────────────────────────
 * 用一份 .xlsx 当**样式捐赠者**,把它的 styling 套到无格式数据源(JSON / CSV)上。
 *
 * **设计本意**:JSON / CSV 这种数据源自身不带格式,如果想让它呈现为漂亮的发票/报表样式,
 * 用户做一份 .xlsx 模板(带表头、边框、配色、列宽、合并、冻结、字体),然后用数据替换模板的文字内容。
 *
 * **语义(必读 — 跟此前 placeholder + anchor 模型彻底不一样)**:
 * - 模板的所有 raw 文字 / 占位符 / 装饰文字 / 单元格内嵌图(DISPIMG) / 浮动图 / 图表 / 条件格式 **全部丢弃**
 * - 模板保留的只有: **styles 池 + 各格 styleId + merges + 列宽 + 行高 + freeze + 网格线开关 + theme 调色板**
 * - 数据(JSON / CSV)的格保留**自然位置 + raw + type**, styleId 取**模板同位置格的 styleId**(没有则 0)
 * - 数据的 sheet 名 / date1904 / cellImages 透传
 *
 * **不适用 .xlsx 数据源** —— .xlsx 自带格式,壳层会忽略 templateFile 并 console.warn.
 * 此函数只负责数据合成,模式判断在调用方.
 */
import type { CellModel, SheetModel, WorkbookModel } from '../model/types'
import { cellKey } from '../model/types'

/**
 * 把模板的样式套到数据 workbook 上,产出新 workbook(不修改入参).
 *
 * @param dataWb     数据源(JSON / CSV 加载得到的 WorkbookModel),raw 值的来源
 * @param templateWb 模板源(.xlsx 解析得到的 WorkbookModel),styling 的来源
 * @returns 合成 workbook —— 用模板的 sheet 形态 + 数据的 raw 内容
 */
export function applyStyleTemplate(
  dataWb: WorkbookModel,
  templateWb: WorkbookModel,
): WorkbookModel {
  const tplSheet = templateWb.sheets[0]
  const dataSheet: SheetModel | undefined = dataWb.sheets[0]
  if (!tplSheet) return dataWb
  if (!dataSheet) {
    // 没数据 → 给个干净的模板克隆(清掉所有 raw,只剩 styling)
    return {
      sheets: [stripTemplateContent(tplSheet, tplSheet.name)],
      activeSheet: 0,
      themeColors: templateWb.themeColors,
      date1904: dataWb.date1904,
      cellImages: dataWb.cellImages,
    }
  }

  // 输出 sheet:模板 sheet 的结构 + 数据 sheet 的内容
  const out: SheetModel = {
    name: dataSheet.name || tplSheet.name,
    index: 0,
    state: 'visible',
    dimension: { rows: 0, cols: 0 },
    cells: new Map<string, CellModel>(),
    styles: tplSheet.styles.map((s) => ({ ...s })),
    merges: tplSheet.merges.map((m) => ({ ...m })),
    columns: new Map(tplSheet.columns),
    rows: new Map(tplSheet.rows),
    defaultColWidth: tplSheet.defaultColWidth,
    defaultRowHeight: tplSheet.defaultRowHeight,
    freeze: { ...tplSheet.freeze },
    // 条件格式 / 数据验证 不带过来 —— 数据可能不在模板的目标列上, 套规则会误命中
    conditional: [],
    dataValidations: [],
    // 模板的图 / 图表 / 形状 全部不带 —— 模板是"样式骨架", 不是内容
    images: [],
    charts: [],
    shapes: [],
    sparklines: [],
    pivotTables: [],
    showGridLines: tplSheet.showGridLines,
  }

  // 数据格按自然位置写入,styleId 从模板同位置取(没模板格 → 0 默认)
  let maxRow = 0
  let maxCol = 0
  for (const cell of dataSheet.cells.values()) {
    const tplCell = tplSheet.cells.get(cellKey(cell.row, cell.col))
    out.cells.set(cellKey(cell.row, cell.col), {
      row: cell.row,
      col: cell.col,
      type: cell.type,
      raw: cell.raw,
      rich: cell.rich,
      formula: cell.formula,
      hyperlink: cell.hyperlink,
      comment: cell.comment,
      styleId: tplCell?.styleId ?? 0,
      dispImgId: cell.dispImgId,
    })
    if (cell.row > maxRow) maxRow = cell.row
    if (cell.col > maxCol) maxCol = cell.col
  }
  // 维度:数据声明 / 实际写入 / 模板列宽行高声明的 cols/rows 三者取大
  // (拿模板列宽是为了让"模板设了 20 列宽但数据只 5 列"的场景也能撑出模板视觉)
  const tplCols = tplSheet.columns.size ? Math.max(...tplSheet.columns.keys()) + 1 : 0
  const tplRows = tplSheet.rows.size ? Math.max(...tplSheet.rows.keys()) + 1 : 0
  out.dimension = {
    rows: Math.max(dataSheet.dimension.rows, maxRow + 1, tplRows),
    cols: Math.max(dataSheet.dimension.cols, maxCol + 1, tplCols),
  }

  return {
    sheets: [out],
    activeSheet: 0,
    themeColors: templateWb.themeColors,
    date1904: dataWb.date1904,
    cellImages: dataWb.cellImages,
  }
}

/** 把模板 sheet 剥成"纯样式骨架"(无 raw / 无图 / 无图表),配合空数据用. */
function stripTemplateContent(tplSheet: SheetModel, name: string): SheetModel {
  return {
    name,
    index: 0,
    state: 'visible',
    dimension: { rows: 0, cols: 0 },
    cells: new Map<string, CellModel>(),
    styles: tplSheet.styles.map((s) => ({ ...s })),
    merges: tplSheet.merges.map((m) => ({ ...m })),
    columns: new Map(tplSheet.columns),
    rows: new Map(tplSheet.rows),
    defaultColWidth: tplSheet.defaultColWidth,
    defaultRowHeight: tplSheet.defaultRowHeight,
    freeze: { ...tplSheet.freeze },
    conditional: [],
    dataValidations: [],
    images: [],
    charts: [],
    shapes: [],
    sparklines: [],
    pivotTables: [],
    showGridLines: tplSheet.showGridLines,
  }
}
