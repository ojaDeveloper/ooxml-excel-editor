/**
 * 工作簿深克隆 / 就地还原(框架无关)—— 脏状态"放弃修改、还原原件"的承重件。
 * 模型是纯数据(Map / Date / Uint8Array / 普通对象),structuredClone 可整体克隆。
 *
 * 关键:`restoreWorkbookInto` **不换对象身份** —— 把 baseline 的内容就地灌回 live 的
 * 各 SheetModel(Object.assign 换字段值,但保留 sheet 对象 + workbook + sheets 数组身份),
 * 于是渲染器、两壳持有的 model 引用在还原后仍然有效。
 */
import type { WorkbookModel } from './types'

/** 深克隆整个工作簿(供 editable 时懒捕获 baseline)。 */
export function cloneWorkbook(wb: WorkbookModel): WorkbookModel {
  return structuredClone(wb)
}

/**
 * 把 baseline(snap)的内容就地还原进 live(不换 live / live.sheets / 各 sheet 的对象身份)。
 * 每次都对 snap 取一份新克隆,故 baseline 自身保持原始、可重复还原。
 * 约定:editing 期间不增删 sheet,故按 index 对齐;多出的 live sheet 不动。
 */
export function restoreWorkbookInto(live: WorkbookModel, snap: WorkbookModel): void {
  const fresh = structuredClone(snap)
  for (let i = 0; i < live.sheets.length && i < fresh.sheets.length; i++) {
    Object.assign(live.sheets[i], fresh.sheets[i]) // 换字段值,保 sheet 对象身份
  }
  live.activeSheet = fresh.activeSheet
  live.date1904 = fresh.date1904
  live.themeColors = fresh.themeColors
}
