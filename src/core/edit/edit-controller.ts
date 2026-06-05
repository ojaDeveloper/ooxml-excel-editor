/**
 * EditController(框架无关)—— 编辑底座:命令栈(undo/redo)、经 mutations 应用命令、
 * 发"前后完整快照"事件、暴露命令式编辑 API + 查询 API。组合进 ViewerController(非继承)。
 * 这是要求 5("一切都有 API/事件")的承重墙:UI/公式/导出都建在它上面。
 */
import type { CellStyleOverride, ColumnInfo, ImageAnchor, MergeRange, RowInfo, SheetModel, WorkbookModel } from '../model/types'
import type { CellValue } from '../model/data-access'
import { buildCellSnapshot, type CellSnapshot } from '../model/snapshot'
import { cloneWorkbook, restoreWorkbookInto } from '../model/clone'
import { cloneImageAnchor } from '../model/mutations'
import { applyCommand, affectedOf, isDimCommand, isImageCommand, isStructCommand, type CellPos, type DimAxis, type EditCommand } from './commands'
import { applyStructOp, type StructOp } from '../model/structure'
import { rewriteWorkbookFormulas } from '../formula/refs'
import type { FormulaEngine, FormulaEngineFactory } from '../formula/engine'
import { collectDirty, writeDirty, dependentsOnSheet } from '../formula/recalc'

export type EditEventName =
  | 'cell-change'
  | 'edit-start'
  | 'edit-commit'
  | 'dim-change'
  | 'dirty-change'
  | 'image-change'
  | 'struct-change'
export type EditSource = 'api' | 'ui' | 'undo' | 'redo'

/** 结构变更事件载荷(增删行列;restore = 撤销/重做的整体还原) */
export interface StructChangePayload {
  op: StructOp | 'restore'
  at?: number
  count?: number
  source: EditSource
}

/** 图片变更事件载荷(before/after = ImageAnchor 克隆;add → before=null;remove → after=null) */
export interface ImageChangePayload {
  index: number
  before: ImageAnchor | null
  after: ImageAnchor | null
  source: EditSource
}

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
  /** 当前活动表索引(0-based;重算把级联事件/写回锚到它) */
  getActiveSheetIndex(): number
  /** 是否开启公式重算(EditConfig.recalc) */
  isRecalcEnabled(): boolean
  /** 公式引擎工厂(EditConfig.formulaEngine,或默认 HyperFormula;未开重算 → null) */
  getEngineFactory(): FormulaEngineFactory | null
  /** 模型变更后回调:重建几何 + 重绘 */
  onModelChange(): void
  /** 图片结构变化(增删)→ 重建叠加层(重绘不够,要新建/移除 DOM)。移动只需 onModelChange 重定位。 */
  rebuildOverlays(): void
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
  // 公式引擎(E4;可选,异步懒初始化;未就绪时编辑不重算)
  private engine: FormulaEngine | null = null
  private engineFor: WorkbookModel | null = null // 引擎对应哪个 workbook
  private warming = false

  constructor(private host: EditControllerHost) {}

  // ---- 公式引擎生命周期(E4) ----
  /** 异步懒初始化引擎(开重算 + 有工厂 + 尚未为当前簿建好)。返回的 Promise 供测试 await;生产 fire-and-forget。 */
  warmEngine(): Promise<void> {
    if (this.warming || !this.host.isRecalcEnabled()) return Promise.resolve()
    const factory = this.host.getEngineFactory()
    const wb = this.host.getWorkbook()
    if (!factory || !wb || this.engineFor === wb) return Promise.resolve()
    this.warming = true
    return (async () => {
      try {
        const eng = await factory()
        if (this.host.getWorkbook() !== wb) {
          eng.destroy() // 等待期间换了簿 → 丢弃
          return
        }
        eng.setSheets(wb)
        this.engine?.destroy()
        this.engine = eng
        this.engineFor = wb
      } catch (e) {
        console.warn('[ooxml-preview] 公式引擎初始化失败(重算已跳过):', e)
      } finally {
        this.warming = false
      }
    })()
  }
  /** 释放引擎(切簿/关重算/dispose)。 */
  disposeEngine(): void {
    this.engine?.destroy()
    this.engine = null
    this.engineFor = null
  }
  /** 配置变化(recalc/factory)或换新簿:重置引擎并按需重新点火。 */
  refreshEngine(): void {
    this.disposeEngine()
    this.warmEngine()
  }
  private engineReady(): boolean {
    return !!this.engine && this.host.isRecalcEnabled() && this.engineFor === this.host.getWorkbook()
  }
  /** 公式引擎是否已就绪(异步 warm 完成且对应当前簿);未开重算恒 false。 */
  isRecalcReady(): boolean {
    return this.engineReady()
  }

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

  /** 给区域套样式覆盖(E5;跳过只读格)。返回是否有改动。前后 style 不同 → 发 cell-change。 */
  setStyle(range: MergeRange, patch: CellStyleOverride): boolean {
    if (!this.host.isEditingEnabled()) return false
    const cells: CellPos[] = []
    for (let r = range.top; r <= range.bottom; r++)
      for (let c = range.left; c <= range.right; c++) if (this.host.isEditable(r, c)) cells.push({ row: r, col: c })
    if (!cells.length) return false
    this.ensureBaseline()
    const inv = this.exec({ kind: 'set-style', cells, patch }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }

  // ---- 行列结构编辑(E7;增删行列,快照逆) ----
  insertRows(at: number, count = 1): boolean {
    return this.structEdit('insert-rows', at, count)
  }
  deleteRows(at: number, count = 1): boolean {
    return this.structEdit('delete-rows', at, count)
  }
  insertCols(at: number, count = 1): boolean {
    return this.structEdit('insert-cols', at, count)
  }
  deleteCols(at: number, count = 1): boolean {
    return this.structEdit('delete-cols', at, count)
  }
  private structEdit(op: StructOp, at: number, count: number): boolean {
    if (!this.host.isEditingEnabled() || at < 0 || count <= 0) return false
    this.ensureBaseline()
    const inv = this.exec({ kind: 'struct-edit', op, at, count }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }

  // ---- 合并单元格(G1) ----
  /** 合并区域(吸收相交旧合并,清空被覆盖格只留左上锚点)。单格不合并。 */
  mergeCells(range: MergeRange): boolean {
    if (!this.host.isEditingEnabled()) return false
    if (range.top === range.bottom && range.left === range.right) return false
    this.ensureBaseline()
    const inv = this.exec({ kind: 'merge-cells', range }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }
  /** 拆分:移除与区域相交的所有合并。 */
  unmergeCells(range: MergeRange): boolean {
    if (!this.host.isEditingEnabled()) return false
    this.ensureBaseline()
    const inv = this.exec({ kind: 'unmerge-cells', range }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }

  // ---- 图片编辑(浮动/嵌入;E6) ----
  /** 读当前表全部图片锚点(克隆,防外部改)。 */
  getImages(): ImageAnchor[] {
    const s = this.host.getSheet()
    return s ? s.images.map(cloneImageAnchor) : []
  }
  /** 加一张图,返回插入索引(失败 -1)。 */
  addImage(anchor: ImageAnchor): number {
    if (!this.host.isEditingEnabled()) return -1
    this.ensureBaseline()
    const inv = this.exec({ kind: 'image-add', anchor }, 'api')
    if (inv && inv.kind === 'image-remove') {
      this.pushUndo(inv)
      this.markDirty()
      return inv.index
    }
    return -1
  }
  /** 删一张图。 */
  removeImage(index: number): boolean {
    if (!this.host.isEditingEnabled()) return false
    const s = this.host.getSheet()
    if (!s || !s.images[index]) return false
    this.ensureBaseline()
    const inv = this.exec({ kind: 'image-remove', index }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }
  /**
   * 补登一次图片移动/缩放(拖拽/programmatic:模型已被 setImageRect 改完,这里只补 undo + 发 image-change)。
   * baseline 须在变更前由调用方 ensureBaseline() 捕获。
   */
  recordImageEdit(index: number, before: ImageAnchor, after: ImageAnchor): void {
    this.pushUndo({ kind: 'image-set', index, anchor: before })
    this.markDirty()
    this.host.emit('image-change', { index, before, after, source: 'ui' } satisfies ImageChangePayload)
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
    // 图片族:发 image-change(增删重建叠加层 / 移动重定位)
    if (isImageCommand(cmd)) {
      const index = cmd.kind === 'image-add' ? (cmd.index ?? sheet.images.length) : cmd.index
      const before = cmd.kind === 'image-add' ? null : (sheet.images[cmd.index] ? cloneImageAnchor(sheet.images[cmd.index]) : null)
      const { inverse } = applyCommand(sheet, cmd)
      const after = cmd.kind === 'image-remove' ? null : (sheet.images[index] ? cloneImageAnchor(sheet.images[index]) : null)
      if (cmd.kind === 'image-set') this.host.onModelChange()
      else this.host.rebuildOverlays()
      this.host.emit('image-change', { index, before, after, source } satisfies ImageChangePayload)
      return inverse
    }
    // 结构族(增删行列):整簿快照逆 + 跨表公式引用重写;全表重建(几何+合并索引+叠加层)
    if (isStructCommand(cmd)) {
      const wb = this.host.getWorkbook()
      if (!wb) return null
      const si = this.host.getActiveSheetIndex()
      let inverse: EditCommand
      if (cmd.kind === 'struct-edit') {
        const snap = cloneWorkbook(wb) // 整簿快照(跨表公式重写也要可撤销)
        applyStructOp(wb.sheets[si], cmd.op, cmd.at, cmd.count)
        rewriteWorkbookFormulas(wb, si, cmd.op, cmd.at, cmd.count) // F1:全簿公式引用重写
        inverse = { kind: 'restore-wb', snapshot: snap }
      } else {
        const cur = cloneWorkbook(wb)
        restoreWorkbookInto(wb, cmd.snapshot)
        inverse = { kind: 'restore-wb', snapshot: cur }
      }
      this.host.onModelChange() // rebuildMetrics(含 merges 索引)+ 重绘
      this.host.rebuildOverlays() // 图片随结构移位/移除
      if (this.host.isRecalcEnabled()) this.refreshEngine() // 引擎按新结构 + 新公式文本重建
      const payload: StructChangePayload =
        cmd.kind === 'struct-edit'
          ? { op: cmd.op, at: cmd.at, count: cmd.count, source }
          : { op: 'restore', source }
      this.host.emit('struct-change', payload)
      return inverse
    }
    // 单元格族:逐格发前后完整快照
    const d = this.host.getDate1904()
    const affected: CellPos[] = affectedOf(cmd)
    const before = affected.map((p) => buildCellSnapshot(sheet, p.row, p.col, d))
    const { inverse } = applyCommand(sheet, cmd)

    // 公式重算(开启 + 引擎就绪):同步进引擎 → 拿级联脏格 → 写回计算值。
    // 依赖格(非直接编辑)在写回前拍前态,写回后拍后态,逐格发 cell-change。
    // set-style 不改内容(只改 styleId)→ 跳过重算,省引擎往返。
    let deps: CellPos[] = []
    let depBefore: CellSnapshot[] = []
    if (cmd.kind !== 'set-style' && this.engineReady()) {
      const wb = this.host.getWorkbook()!
      const si = this.host.getActiveSheetIndex()
      const dirty = collectDirty(this.engine!, wb, si, affected)
      deps = dependentsOnSheet(dirty, si, affected)
      depBefore = deps.map((p) => buildCellSnapshot(sheet, p.row, p.col, d))
      writeDirty(wb, dirty) // 写回(含编辑的公式格自身 + 依赖格)
    }

    this.host.onModelChange()
    affected.forEach((p, i) => {
      const after = buildCellSnapshot(sheet, p.row, p.col, d)
      this.host.emit('cell-change', { before: before[i], after, source } satisfies CellChangePayload)
    })
    deps.forEach((p, i) => {
      const after = buildCellSnapshot(sheet, p.row, p.col, d)
      this.host.emit('cell-change', { before: depBefore[i], after, source } satisfies CellChangePayload)
    })
    return inverse
  }
}
