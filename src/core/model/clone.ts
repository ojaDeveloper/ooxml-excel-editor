/**
 * 工作簿轻量克隆 / 就地还原(框架无关)—— 脏状态"还原原件"+ 结构编辑(增删行列)undo 的承重件。
 *
 * 性能要点:**只克隆编辑会改的部分**(cells/merges/行列/dimension/freeze/styles 数组),
 * **不可变的重负载按引用共享**(图片字节、图表/形状/迷你图/条件格式/数据验证)。编辑期间从不
 * 原地改这些共享部分(结构编辑只挪 image 锚点对象、不碰 bytes/charts/pivotTables),所以共享安全 ——
 * 避免了 structuredClone 把几 MB 图片字节、图表一并深拷,大文件 + undo 栈内存大幅下降。
 *
 * `restoreWorkbookInto` **不换对象身份** —— 把内容就地灌回 live 的各 SheetModel(保 sheet 对象 +
 * workbook + sheets 数组身份),渲染器 / 两壳持有的 model 引用在还原后仍有效。
 */
import type { SheetModel, WorkbookModel } from './types'
import { cloneCell } from './snapshot'
import { cloneImageAnchor } from './mutations'

/** 轻量克隆一张表:编辑会动的字段深克隆,其余(charts/shapes/conditional/DV/pivotTables…)按引用共享。 */
function cloneSheet(s: SheetModel): SheetModel {
  return {
    ...s, // 共享不可变重负载:charts/shapes/sparklines/conditional/dataValidations/pageSetup/name…
    cells: new Map(Array.from(s.cells, ([k, c]) => [k, cloneCell(c)])),
    merges: s.merges.map((m) => ({ ...m })),
    rows: new Map(Array.from(s.rows, ([k, v]) => [k, { ...v }])),
    columns: new Map(Array.from(s.columns, ([k, v]) => [k, { ...v }])),
    images: s.images.map(cloneImageAnchor), // 新锚点对象、共享 bytes(不可变)
    dimension: { ...s.dimension },
    freeze: { ...s.freeze },
    styles: s.styles.slice(), // 新数组(防 internStyle 追加污染快照);style 对象共享(编辑不原地改)
    pivotTables: (s.pivotTables ?? []).map((p) => ({
      ...p,
      range: { ...p.range },
      fields: p.fields.slice(),
      buttons: p.buttons.map((b) => ({ ...b })),
      source: p.source ? { sheetIndex: p.source.sheetIndex, range: { ...p.source.range } } : undefined,
      layout: p.layout ? {
        filters: p.layout.filters.map((rule) => ({ ...rule, values: rule.values?.slice() })),
        columns: p.layout.columns.slice(),
        rows: p.layout.rows.slice(),
        values: p.layout.values.map((rule) => ({ ...rule })),
      } : undefined,
      collapsed: p.collapsed?.slice(),
      rowGroups: p.rowGroups?.map((g) => ({ ...g })),
    })),
  }
}

/** 轻量深克隆整个工作簿(供 editable 懒捕获 baseline + 结构编辑 undo 快照)。 */
export function cloneWorkbook(wb: WorkbookModel): WorkbookModel {
  return {
    ...wb,
    sheets: wb.sheets.map(cloneSheet),
    themeColors: wb.themeColors.slice(),
    // WPS 内嵌图登记表:新 Map(防转换增删污染快照),CellImage 不可变 → 按引用共享
    cellImages: wb.cellImages ? new Map(wb.cellImages) : undefined,
  }
}

/**
 * 把 baseline/快照(snap)的内容就地还原进 live(不换 live / live.sheets / 各 sheet 的对象身份)。
 * 每次对 snap 取一份新轻量克隆,故 snap 自身保持可重复还原(redo / 多次 reset)。
 */
export function restoreWorkbookInto(live: WorkbookModel, snap: WorkbookModel): void {
  const fresh = cloneWorkbook(snap)
  for (let i = 0; i < live.sheets.length && i < fresh.sheets.length; i++) {
    Object.assign(live.sheets[i], fresh.sheets[i]) // 换字段值,保 sheet 对象身份
  }
  live.sheets.splice(fresh.sheets.length)
  for (let i = live.sheets.length; i < fresh.sheets.length; i++) live.sheets.push(fresh.sheets[i])
  live.activeSheet = fresh.activeSheet
  live.date1904 = fresh.date1904
  live.themeColors = fresh.themeColors
  live.cellImages = fresh.cellImages // 转换 undo/redo 要还原登记表
}
