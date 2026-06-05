/**
 * 重算编排(框架无关)—— 一次编辑后把改动同步进引擎、取依赖级联脏格、把计算结果写回模型。
 * 公式文本在 cell.formula、结果写回 cell.raw —— 现有读契约(cellDisplayText/导出)零改。
 *
 * 拆成"收集脏格 / 写回"两步:调用方(EditController.exec)在写回前先拍依赖格的前态快照,
 * 这样 cell-change 的 before/after 才准确。undo/redo 也走这里(restore-cells 后再同步引擎+重算),
 * 故依赖格永远由引擎推导、不进命令载荷 —— 命令只携带"直接编辑的格",级联恒定可复算(确定性)。
 */
import type { WorkbookModel } from '../model/types'
import { cellKey } from '../model/types'
import { setCellComputed } from '../model/mutations'
import { cellContentForEngine, type DirtyCell, type FormulaEngine } from './engine'
import type { CellPos } from '../edit/commands'

/** 把 editedCells 同步进引擎,收集所有受影响格(含跨表依赖),**不写模型**。 */
export function collectDirty(
  engine: FormulaEngine,
  wb: WorkbookModel,
  activeSheet: number,
  editedCells: CellPos[],
): DirtyCell[] {
  const sheetModel = wb.sheets[activeSheet]
  const dirty = new Map<string, DirtyCell>()
  for (const p of editedCells) {
    const cell = sheetModel?.cells.get(cellKey(p.row, p.col)) ?? null
    for (const dc of engine.setCell(activeSheet, p.row, p.col, cellContentForEngine(cell))) {
      dirty.set(`${dc.sheet}:${dc.row}:${dc.col}`, dc)
    }
  }
  return [...dirty.values()]
}

/** 把脏格的计算值写回各自模型表(只改 raw)。 */
export function writeDirty(wb: WorkbookModel, dirty: DirtyCell[]): void {
  for (const dc of dirty) {
    const tgt = wb.sheets[dc.sheet]
    if (tgt) setCellComputed(tgt, dc.row, dc.col, dc.value)
  }
}

/** 活动表上、由级联产生(非直接编辑)的脏格 —— 供逐格发 cell-change。 */
export function dependentsOnSheet(dirty: DirtyCell[], activeSheet: number, editedCells: CellPos[]): CellPos[] {
  const edited = new Set(editedCells.map((p) => cellKey(p.row, p.col)))
  return dirty
    .filter((d) => d.sheet === activeSheet && !edited.has(cellKey(d.row, d.col)))
    .map((d) => ({ row: d.row, col: d.col }))
}
