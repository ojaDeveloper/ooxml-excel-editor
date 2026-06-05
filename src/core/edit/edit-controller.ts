/**
 * EditController(框架无关)—— 编辑底座:命令栈(undo/redo)、经 mutations 应用命令、
 * 发"前后完整快照"事件、暴露命令式编辑 API + 查询 API。组合进 ViewerController(非继承)。
 * 这是要求 5("一切都有 API/事件")的承重墙:UI/公式/导出都建在它上面。
 */
import type { MergeRange, SheetModel } from '../model/types'
import type { CellValue } from '../model/data-access'
import { buildCellSnapshot, type CellSnapshot } from '../model/snapshot'
import { applyCommand, affectedOf, type CellPos, type EditCommand } from './commands'

export type EditEventName = 'cell-change' | 'edit-start' | 'edit-commit'
export type EditSource = 'api' | 'ui' | 'undo' | 'redo'

export interface CellChangePayload {
  before: CellSnapshot
  after: CellSnapshot
  source: EditSource
}

/** EditController 与 ViewerController 的桥(host-callback,框架无关) */
export interface EditControllerHost {
  getSheet(): SheetModel | null
  getDate1904(): boolean
  /** 该格是否可编辑(综合 editable + readOnlyRanges + cellReadOnly) */
  isEditable(row: number, col: number): boolean
  /** 模型变更后回调:重建几何 + 重绘 */
  onModelChange(): void
  /** 发编辑事件(壳转 emit + 插件派发) */
  emit(event: EditEventName, payload: unknown): void
}

export class EditController {
  private undoStack: EditCommand[] = []
  private redoStack: EditCommand[] = []
  private editing: CellPos | null = null
  private readonly limit = 200

  constructor(private host: EditControllerHost) {}

  // ---- 状态查询 ----
  getEditingCell(): CellPos | null {
    return this.editing
  }
  setEditing(cell: CellPos | null): void {
    this.editing = cell
  }
  canUndo(): boolean {
    return this.undoStack.length > 0
  }
  canRedo(): boolean {
    return this.redoStack.length > 0
  }
  /** 查询任意格的完整快照(要求 5:事件之外也有查询 API) */
  getCellSnapshot(row: number, col: number): CellSnapshot | null {
    const s = this.host.getSheet()
    return s ? buildCellSnapshot(s, row, col, this.host.getDate1904()) : null
  }
  /** 切表/换簿时清空(交互态 + 命令栈作废) */
  reset(): void {
    this.undoStack = []
    this.redoStack = []
    this.editing = null
  }

  // ---- 命令式编辑 API(要求 3:直接编辑哪行哪列) ----
  /** 编辑单格;只读则不动,返回是否生效 */
  editCell(row: number, col: number, value: CellValue): boolean {
    if (!this.host.isEditable(row, col)) return false
    const inv = this.exec({ kind: 'set-value', row, col, value }, 'api')
    if (inv) this.pushUndo(inv)
    return !!inv
  }
  /** 区域批量设值(2D,左上对齐 range.top/left);跳过只读格,返回是否有改动 */
  editRange(range: MergeRange, values: CellValue[][]): boolean {
    const cells = this.collectEditable(range, values)
    if (!cells.length) return false
    const inv = this.exec({ kind: 'set-cells', cells }, 'api')
    if (inv) this.pushUndo(inv)
    return !!inv
  }
  /** 清空区域(跳过只读) */
  clearRange(range: MergeRange): boolean {
    const cells: { row: number; col: number; value: CellValue }[] = []
    for (let r = range.top; r <= range.bottom; r++)
      for (let c = range.left; c <= range.right; c++) if (this.host.isEditable(r, c)) cells.push({ row: r, col: c, value: null })
    if (!cells.length) return false
    const inv = this.exec({ kind: 'set-cells', cells }, 'api')
    if (inv) this.pushUndo(inv)
    return !!inv
  }

  undo(): void {
    const cmd = this.undoStack.pop()
    if (!cmd) return
    const inv = this.exec(cmd, 'undo')
    if (inv) this.redoStack.push(inv)
  }
  redo(): void {
    const cmd = this.redoStack.pop()
    if (!cmd) return
    const inv = this.exec(cmd, 'redo')
    if (inv) this.undoStack.push(inv)
  }

  // ---- 内部 ----
  private collectEditable(range: MergeRange, values: CellValue[][]): { row: number; col: number; value: CellValue }[] {
    const out: { row: number; col: number; value: CellValue }[] = []
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const row = range.top + r
        const col = range.left + c
        if (this.host.isEditable(row, col)) out.push({ row, col, value: values[r][c] })
      }
    }
    return out
  }

  private pushUndo(inverse: EditCommand): void {
    this.undoStack.push(inverse)
    if (this.undoStack.length > this.limit) this.undoStack.shift()
    this.redoStack = [] // 新编辑使 redo 作废
  }

  /** 应用一条命令:建前快照 → apply → 重绘 → 逐格发后快照事件。返回逆命令。 */
  private exec(cmd: EditCommand, source: EditSource): EditCommand | null {
    const sheet = this.host.getSheet()
    if (!sheet) return null
    const d = this.host.getDate1904()
    const affected: CellPos[] = affectedOf(cmd)
    const before = affected.map((p) => buildCellSnapshot(sheet, p.row, p.col, d))
    const { inverse } = applyCommand(sheet, cmd)
    this.host.onModelChange()
    affected.forEach((p, i) => {
      const after = buildCellSnapshot(sheet, p.row, p.col, d)
      this.host.emit('cell-change', { before: before[i], after, source } satisfies CellChangePayload)
    })
    return inverse
  }
}
