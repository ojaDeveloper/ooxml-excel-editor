/**
 * EditController(框架无关)—— 编辑底座:命令栈(undo/redo)、经 mutations 应用命令、
 * 发"前后完整快照"事件、暴露命令式编辑 API + 查询 API。组合进 ViewerController(非继承)。
 * 这是要求 5("一切都有 API/事件")的承重墙:UI/公式/导出都建在它上面。
 */
import type { ColumnInfo, MergeRange, RowInfo, SheetModel, WorkbookModel } from '../model/types'
import type { CellValue } from '../model/data-access'
import { buildCellSnapshot, type CellSnapshot } from '../model/snapshot'
import { cloneWorkbook, restoreWorkbookInto } from '../model/clone'
import { applyCommand, affectedOf, isDimCommand, type CellPos, type DimAxis, type EditCommand } from './commands'

export type EditEventName = 'cell-change' | 'edit-start' | 'edit-commit' | 'dim-change' | 'dirty-change'
export type EditSource = 'api' | 'ui' | 'undo' | 'redo'

export interface CellChangePayload {
  before: CellSnapshot
  after: CellSnapshot
  source: EditSource
}

/** 列宽/行高变更事件载荷(before/after = px 尺寸,含默认值回落) */
export interface DimChangePayload {
  axis: DimAxis
  index: number
  before: number
  after: number
  source: EditSource
}

/** 脏状态变更事件载荷 */
export interface DirtyChangePayload {
  dirty: boolean
}

/** EditController 与 ViewerController 的桥(host-callback,框架无关) */
export interface EditControllerHost {
  getSheet(): SheetModel | null
  getWorkbook(): WorkbookModel | null
  getDate1904(): boolean
  /** 该格是否可编辑(综合 editable + readOnlyRanges + cellReadOnly) */
  isEditable(row: number, col: number): boolean
  /** 编辑模式是否开启(决定是否懒捕获 baseline / 记账脏状态) */
  isEditingEnabled(): boolean
  /** 模型变更后回调:重建几何 + 重绘 */
  onModelChange(): void
  /** 发编辑事件(壳转 emit + 插件派发) */
  emit(event: EditEventName, payload: unknown): void
}

/** 取某列/行当前 px 尺寸(无 Map 项 → 回落默认宽高) */
function currentDimSize(sheet: SheetModel, axis: DimAxis, index: number): number {
  if (axis === 'col') return sheet.columns.get(index)?.width ?? sheet.defaultColWidth
  return sheet.rows.get(index)?.height ?? sheet.defaultRowHeight
}

export class EditController {
  private undoStack: EditCommand[] = []
  private redoStack: EditCommand[] = []
  private editing: CellPos | null = null
  private readonly limit = 200
  // 脏状态 + 原件 baseline(均 workbook 级;editable 时懒捕获,读层零成本)
  private dirty = false
  private baseline: WorkbookModel | null = null
  private baselineFor: WorkbookModel | null = null

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
    this.ensureBaseline()
    const inv = this.exec({ kind: 'set-value', row, col, value }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }
  /** 区域批量设值(2D,左上对齐 range.top/left);跳过只读格,返回是否有改动 */
  editRange(range: MergeRange, values: CellValue[][]): boolean {
    const cells = this.collectEditable(range, values)
    if (!cells.length) return false
    this.ensureBaseline()
    const inv = this.exec({ kind: 'set-cells', cells }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }
  /** 清空区域(跳过只读) */
  clearRange(range: MergeRange): boolean {
    const cells: { row: number; col: number; value: CellValue }[] = []
    for (let r = range.top; r <= range.bottom; r++)
      for (let c = range.left; c <= range.right; c++) if (this.host.isEditable(r, c)) cells.push({ row: r, col: c, value: null })
    if (!cells.length) return false
    this.ensureBaseline()
    const inv = this.exec({ kind: 'set-cells', cells }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }

  // ---- 维度编辑(列宽/行高;E3.5:resize 入命令栈) ----
  /** 程序化设列宽/行高(API 路径:apply-via-command)。返回是否生效。 */
  setDimension(axis: DimAxis, index: number, size: number): boolean {
    if (!this.host.isEditingEnabled()) return false
    this.ensureBaseline()
    const inv = this.exec({ kind: 'set-dim', axis, index, size }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }
  /**
   * 补登一次维度变更(拖拽/autofit 路径:模型已被 renderer 改完,这里只补 undo 项 + 发事件)。
   * baseline 须在变更前(拖拽起始/autofit 前)由调用方 ensureBaseline() 捕获。
   */
  recordDimEdit(axis: DimAxis, index: number, beforeInfo: ColumnInfo | RowInfo | null, before: number, after: number): void {
    if (before === after) return
    this.pushUndo({ kind: 'restore-dim', axis, index, info: beforeInfo })
    this.markDirty()
    this.host.emit('dim-change', { axis, index, before, after, source: 'ui' } satisfies DimChangePayload)
  }

  // ---- 脏状态 + 原件还原(E3.5) ----
  /** editable 时懒捕获 baseline(首次编辑/resize 前调,模型仍原始);幂等。 */
  ensureBaseline(): void {
    if (!this.host.isEditingEnabled()) return
    const wb = this.host.getWorkbook()
    if (!wb || this.baselineFor === wb) return
    this.baseline = cloneWorkbook(wb)
    this.baselineFor = wb
  }
  isDirty(): boolean {
    return this.dirty
  }
  private markDirty(): void {
    if (this.dirty) return
    this.dirty = true
    this.host.emit('dirty-change', { dirty: true } satisfies DirtyChangePayload)
  }
  /** 换新工作簿(非切表)时:作废 baseline + 清脏(切表保留,见 controller 判定)。 */
  resetDirtyBaseline(): void {
    this.baseline = null
    this.baselineFor = null
    if (this.dirty) {
      this.dirty = false
      this.host.emit('dirty-change', { dirty: false } satisfies DirtyChangePayload)
    }
  }
  /** 放弃全部修改,还原到刚加载的原件。返回是否还原(无 baseline → false)。 */
  resetToOriginal(): boolean {
    const wb = this.host.getWorkbook()
    if (!wb || !this.baseline) return false
    restoreWorkbookInto(wb, this.baseline)
    this.undoStack = []
    this.redoStack = []
    this.editing = null
    this.host.onModelChange()
    if (this.dirty) {
      this.dirty = false
      this.host.emit('dirty-change', { dirty: false } satisfies DirtyChangePayload)
    }
    return true
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

  /** 应用一条命令:建前快照 → apply → 重绘 → 发事件(cell 族 cell-change / dim 族 dim-change)。返回逆命令。 */
  private exec(cmd: EditCommand, source: EditSource): EditCommand | null {
    const sheet = this.host.getSheet()
    if (!sheet) return null
    // 维度族:发 dim-change(before/after = px 尺寸)
    if (isDimCommand(cmd)) {
      const before = currentDimSize(sheet, cmd.axis, cmd.index)
      const { inverse } = applyCommand(sheet, cmd)
      this.host.onModelChange()
      const after = currentDimSize(sheet, cmd.axis, cmd.index)
      this.host.emit('dim-change', { axis: cmd.axis, index: cmd.index, before, after, source } satisfies DimChangePayload)
      return inverse
    }
    // 单元格族:逐格发前后完整快照
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
