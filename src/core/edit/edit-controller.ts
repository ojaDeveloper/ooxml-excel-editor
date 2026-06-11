/**
 * EditController(框架无关)—— 编辑底座:命令栈(undo/redo)、经 mutations 应用命令、
 * 发"前后完整快照"事件、暴露命令式编辑 API + 查询 API。组合进 ViewerController(非继承)。
 * 这是要求 5("一切都有 API/事件")的承重墙:UI/公式/导出都建在它上面。
 */
import type { CellModel, CellStyleOverride, ColumnInfo, ImageAnchor, MergeRange, RowInfo, SheetModel, WorkbookModel } from '../model/types'
import { cellKey } from '../model/types'
import type { CellValue } from '../model/data-access'
import { buildCellSnapshot, type CellSnapshot } from '../model/snapshot'
import { cloneWorkbook, restoreWorkbookInto } from '../model/clone'
import { cloneImageAnchor, convertFloatToCellImage, convertCellImageToFloat, setCellValue, setCellModel, applyStyleOverride, addImage, internStyle, restoreDimension } from '../model/mutations'
import { pxToEmu } from '../layout/units'
import type { ParsedClipboard } from './clipboard-html'
import { reviveClipRaw, b64ToBytes, type ClipSnapshot } from './clipboard-snapshot'
import { applyCommand, affectedOf, isDimCommand, isImageCommand, isStructCommand, type CellPos, type DimAxis, type EditCommand } from './commands'
import type { PermissionDeniedPayload } from '../plugin'

/** 区域是否全可编辑(per-cell 检查;走 host.isEditable 闸门). 任一只读 → 返 false. */
function _rangeAllEditableViaHost(host: EditControllerHost, range: MergeRange): boolean {
  for (let r = range.top; r <= range.bottom; r++) {
    for (let c = range.left; c <= range.right; c++) {
      if (!host.isEditable(r, c)) return false
    }
  }
  return true
}
/** 收集区域里全部不可编辑的格(给 permission-denied payload). */
function _collectDeniedInRangeViaHost(host: EditControllerHost, range: MergeRange): Array<{ row: number; col: number }> {
  const out: Array<{ row: number; col: number }> = []
  for (let r = range.top; r <= range.bottom; r++) {
    for (let c = range.left; c <= range.right; c++) {
      if (!host.isEditable(r, c)) out.push({ row: r, col: c })
    }
  }
  return out
}
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
  | 'permission-denied'
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

/** data:image/png;base64,xxx → {bytes,mime}(粘贴 data-uri 图用);非法返 null */
function dataUrlToBytes(dataUrl: string): { bytes: Uint8Array; mime: string } | null {
  const m = /^data:([^;,]+)(;base64)?,(.*)$/s.exec(dataUrl)
  if (!m || typeof atob === 'undefined') return null
  try {
    const bin = atob(m[3])
    const bytes = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
    return { bytes, mime: m[1] || 'image/png' }
  } catch {
    return null
  }
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
  /**
   * 富粘贴(Excel/WPS 复制的 HTML 解析后):值 + 样式 + 合并 + 图片,**整体单次撤销**。
   * start = 落点左上角。跳过只读格。一次 cloneWorkbook 快照 → 应用全部 → 压一条 restore-wb 逆。
   */
  pasteRich(start: { row: number; col: number }, parsed: ParsedClipboard): boolean {
    if (!this.host.isEditingEnabled()) return false
    const sheet = this.host.getSheet()
    const wb = this.host.getWorkbook()
    if (!sheet || !wb) return false
    const { values, styles, merges, images } = parsed
    if (!values.length && !styles.length && !merges.length && !images.length) return false
    this.ensureBaseline()
    const d = this.host.getDate1904()
    const snap = cloneWorkbook(wb)

    // 受影响格(值 + 样式覆盖)→ 先拍前态,应用后拍后态发 cell-change
    const affectedKeys = new Set<string>()
    const affected: CellPos[] = []
    const mark = (row: number, col: number) => {
      const k = `${row}:${col}`
      if (!affectedKeys.has(k)) {
        affectedKeys.add(k)
        affected.push({ row, col })
      }
    }
    for (let r = 0; r < values.length; r++)
      for (let c = 0; c < values[r].length; c++) {
        const row = start.row + r
        const col = start.col + c
        if (this.host.isEditable(row, col)) mark(row, col)
      }
    for (const s of styles) {
      const row = start.row + s.row
      const col = start.col + s.col
      if (this.host.isEditable(row, col)) mark(row, col)
    }
    const before = affected.map((p) => buildCellSnapshot(sheet, p.row, p.col, d))

    // 应用:值 → 样式 → 合并 → 图片
    for (let r = 0; r < values.length; r++)
      for (let c = 0; c < values[r].length; c++) {
        const row = start.row + r
        const col = start.col + c
        if (this.host.isEditable(row, col)) setCellValue(sheet, row, col, values[r][c])
      }
    for (const s of styles) {
      const row = start.row + s.row
      const col = start.col + s.col
      if (this.host.isEditable(row, col)) applyStyleOverride(sheet, row, col, s.patch)
    }
    // Phase A 补漏 (2026-06-08): 合并 / 图片粘贴 per-cell editable 检查; 跳过的位置 emit permission-denied
    const denied: Array<{ row: number; col: number }> = []
    for (const m of merges) {
      const range = { top: start.row + m.top, left: start.col + m.left, bottom: start.row + m.bottom, right: start.col + m.right }
      if (range.top === range.bottom && range.left === range.right) continue
      // 合并区域内任一格只读 → 跳过这块 + 收集 denied
      if (!_rangeAllEditableViaHost(this.host, range)) {
        denied.push(..._collectDeniedInRangeViaHost(this.host, range))
        continue
      }
      applyCommand(sheet, { kind: 'merge-cells', range })
    }
    for (const im of images) {
      const row = start.row + im.row
      const col = start.col + im.col
      // 目标格只读 → 跳过 + 收集 denied
      if (!this.host.isEditable(row, col)) { denied.push({ row, col }); continue }
      const dec = dataUrlToBytes(im.dataUrl)
      if (!dec) continue
      const idx = addImage(sheet, { src: im.dataUrl, bytes: dec.bytes, mime: dec.mime, from: { col, row, colOffEmu: 0, rowOffEmu: 0 }, extWidthEmu: pxToEmu(96), extHeightEmu: pxToEmu(96) })
      convertFloatToCellImage(wb, sheet, idx, row, col) // 落进单元格(像 WPS 嵌入)
    }

    this.pushUndo({ kind: 'restore-wb', snapshot: snap }) // 整体单次撤销
    this.markDirty()
    this.host.onModelChange()
    this.host.rebuildOverlays() // 图片增删 → 重建叠加层
    if (this.host.isRecalcEnabled()) this.refreshEngine()
    affected.forEach((p, i) =>
      this.host.emit('cell-change', { before: before[i], after: buildCellSnapshot(sheet, p.row, p.col, d), source: 'api' } satisfies CellChangePayload),
    )
    // Phase A: 合并 / 图片粘贴撞只读 → 整次结束发一次 permission-denied
    if (denied.length) {
      const payload: PermissionDeniedPayload = { reason: 'paste', cells: denied, message: `粘贴撞 ${denied.length} 个只读格,已跳过` }
      this.host.emit('permission-denied', payload)
    }
    return true
  }

  /**
   * 1:1 保真粘贴:用复制时序列化进剪贴板的完整模型快照(本组件自己复制的,跨实例 Vue3/Vue2/React 通用)
   * 覆盖式落到目标区 —— 完整还原 值/原始类型/数字格式/全部样式(含边框)/合并/图片(浮动 + DISPIMG)/行高列宽。
   * 跟 pasteRich(外部 HTML 近似)同走命令栈(整体单次撤销)+ 前后快照 cell-change + 只读 permission-denied。
   */
  pasteSnapshot(start: { row: number; col: number }, snap: ClipSnapshot): boolean {
    if (!this.host.isEditingEnabled()) return false
    const sheet = this.host.getSheet()
    const wb = this.host.getWorkbook()
    if (!sheet || !wb) return false
    if (!snap.cells.length && !snap.merges.length && !snap.images.length && !snap.rowHeights.length && !snap.colWidths.length) return false
    this.ensureBaseline()
    const d = this.host.getDate1904()
    const snapWb = cloneWorkbook(wb)

    // 受影响 = 整个粘贴矩形(覆盖式:源里的空格也清掉,保证 1:1)
    const affected: CellPos[] = []
    for (let r = 0; r < snap.rows; r++)
      for (let c = 0; c < snap.cols; c++) {
        const row = start.row + r
        const col = start.col + c
        if (this.host.isEditable(row, col)) affected.push({ row, col })
      }
    const before = affected.map((p) => buildCellSnapshot(sheet, p.row, p.col, d))
    const denied: Array<{ row: number; col: number }> = []

    // 1) 清目标矩形(可编辑格)→ 覆盖式
    for (const p of affected) sheet.cells.delete(cellKey(p.row, p.col))

    // 2) DISPIMG 单元格图字节登记进目标 cellImages(缺则补,已存在复用 id;data-url 兜底渲染)
    if (snap.cellImages.length) {
      if (!wb.cellImages) wb.cellImages = new Map()
      for (const ci of snap.cellImages) {
        if (!wb.cellImages.has(ci.id)) wb.cellImages.set(ci.id, { id: ci.id, bytes: b64ToBytes(ci.b64), mime: ci.mime, src: `data:${ci.mime};base64,${ci.b64}` })
      }
    }

    // 3) 落格(完整模型 + intern 样式到目标表 → 跨表/跨实例 styleId 正确)
    for (const cc of snap.cells) {
      const row = start.row + cc.r
      const col = start.col + cc.c
      if (!this.host.isEditable(row, col)) { denied.push({ row, col }); continue }
      const model: CellModel = { row, col, type: cc.type, raw: reviveClipRaw(cc.raw), styleId: internStyle(sheet, cc.style) }
      if (cc.formula) model.formula = cc.formula
      if (cc.hyperlink) model.hyperlink = cc.hyperlink
      if (cc.comment) model.comment = cc.comment
      if (cc.rich) model.rich = cc.rich
      if (cc.dispImgId) model.dispImgId = cc.dispImgId
      setCellModel(sheet, model)
    }

    // 4) 合并:先去掉目标矩形内已有合并,再按快照重建(只读区跳过 + denied)
    const tr = { top: start.row, left: start.col, bottom: start.row + snap.rows - 1, right: start.col + snap.cols - 1 }
    sheet.merges = sheet.merges.filter((m) => !(m.top >= tr.top && m.left >= tr.left && m.bottom <= tr.bottom && m.right <= tr.right))
    for (const m of snap.merges) {
      const range = { top: start.row + m.top, left: start.col + m.left, bottom: start.row + m.bottom, right: start.col + m.right }
      if (range.top === range.bottom && range.left === range.right) continue
      if (!_rangeAllEditableViaHost(this.host, range)) { denied.push(..._collectDeniedInRangeViaHost(this.host, range)); continue }
      applyCommand(sheet, { kind: 'merge-cells', range })
    }

    // 5) 浮动图(相对锚点 → 目标起点偏移;DISPIMG 已在第 3 步随格落地,这里只处理真浮动图)
    for (const im of snap.images) {
      const fromRow = start.row + im.from.row
      const fromCol = start.col + im.from.col
      if (!this.host.isEditable(fromRow, fromCol)) { denied.push({ row: fromRow, col: fromCol }); continue }
      addImage(sheet, {
        src: `data:${im.mime};base64,${im.b64}`,
        bytes: b64ToBytes(im.b64),
        mime: im.mime,
        from: { col: fromCol, row: fromRow, colOffEmu: im.from.colOffEmu, rowOffEmu: im.from.rowOffEmu },
        to: im.to ? { col: start.col + im.to.col, row: start.row + im.to.row, colOffEmu: im.to.colOffEmu, rowOffEmu: im.to.rowOffEmu } : undefined,
        extWidthEmu: im.extWidthEmu,
        extHeightEmu: im.extHeightEmu,
        editAs: im.editAs,
      })
    }

    // 6) 行高列宽(1:1:源定制过的行/列尺寸搬到目标对应行列)
    for (const dim of snap.rowHeights) restoreDimension(sheet, 'row', start.row + dim.i, { height: dim.height ?? sheet.defaultRowHeight, hidden: !!dim.hidden, customHeight: !!dim.custom })
    for (const dim of snap.colWidths) restoreDimension(sheet, 'col', start.col + dim.i, { width: dim.width ?? sheet.defaultColWidth, hidden: !!dim.hidden })

    this.pushUndo({ kind: 'restore-wb', snapshot: snapWb }) // 整体单次撤销
    this.markDirty()
    this.host.onModelChange()
    this.host.rebuildOverlays() // 图片增删 → 重建叠加层
    if (this.host.isRecalcEnabled()) this.refreshEngine()
    affected.forEach((p, i) =>
      this.host.emit('cell-change', { before: before[i], after: buildCellSnapshot(sheet, p.row, p.col, d), source: 'api' } satisfies CellChangePayload),
    )
    if (denied.length) {
      this.host.emit('permission-denied', { reason: 'paste', cells: denied, message: `粘贴撞 ${denied.length} 个只读格,已跳过` } satisfies PermissionDeniedPayload)
    }
    return true
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
  /** 合并区域(吸收相交旧合并,清空被覆盖格只留左上锚点)。单格不合并。
   *  Phase A 补漏 (2026-06-08): 区域内任一格只读 → 拒绝整次合并 + emit permission-denied. */
  mergeCells(range: MergeRange): boolean {
    if (!this.host.isEditingEnabled()) return false
    if (range.top === range.bottom && range.left === range.right) return false
    const denied = _collectDeniedInRangeViaHost(this.host, range)
    if (denied.length) {
      this.host.emit('permission-denied', { reason: 'merge', cells: denied, message: `合并区域内有 ${denied.length} 个只读格,已拒绝` } satisfies PermissionDeniedPayload)
      return false
    }
    this.ensureBaseline()
    const inv = this.exec({ kind: 'merge-cells', range }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }
  /** 拆分:移除与区域相交的所有合并。
   *  Phase A 补漏 (2026-06-08): 区域内任一格只读 → 拒绝整次拆分 + emit permission-denied. */
  unmergeCells(range: MergeRange): boolean {
    if (!this.host.isEditingEnabled()) return false
    const denied = _collectDeniedInRangeViaHost(this.host, range)
    if (denied.length) {
      this.host.emit('permission-denied', { reason: 'unmerge', cells: denied, message: `拆分区域内有 ${denied.length} 个只读格,已拒绝` } satisfies PermissionDeniedPayload)
      return false
    }
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

  // ---- WPS 内嵌图 ⇄ 浮动图互转(第二期) ----
  /**
   * 浮动图 → 单元格内嵌图(DISPIMG):把第 imageIndex 张浮动图塞进 (row,col) 格。
   * 图缺字节不可转 → 返 false。入命令栈(undo 走整簿快照还原)。
   */
  convertImageToCell(imageIndex: number, row: number, col: number): boolean {
    if (!this.host.isEditingEnabled()) return false
    // Phase A 补漏: 目标格只读 → 拒绝
    if (!this.host.isEditable(row, col)) {
      this.host.emit('permission-denied', { reason: 'image-convert', cells: [{ row, col }], message: '目标格只读,无法嵌入图片' } satisfies PermissionDeniedPayload)
      return false
    }
    this.ensureBaseline()
    const inv = this.exec({ kind: 'convert-to-cell', imageIndex, row, col }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }
  /**
   * 批量把多张浮动图嵌入各自的目标格(整表/整列);一次进撤销栈。返回成功嵌入的张数。
   * targets 的 imageIndex 是调用前的浮动图索引(内部按降序 splice,索引互不干扰)。
   * Phase A 补漏: 过滤掉目标格只读的 targets, denied 一次性 emit.
   */
  convertImagesToCells(targets: { imageIndex: number; row: number; col: number }[]): number {
    if (!this.host.isEditingEnabled() || !targets.length) return 0
    const denied: Array<{ row: number; col: number }> = []
    const allowed: typeof targets = []
    for (const t of targets) {
      if (this.host.isEditable(t.row, t.col)) allowed.push(t)
      else denied.push({ row: t.row, col: t.col })
    }
    if (denied.length) this.host.emit('permission-denied', { reason: 'image-convert', cells: denied, message: `${denied.length} 个目标格只读,已跳过` } satisfies PermissionDeniedPayload)
    if (!allowed.length) return 0
    const sheet = this.host.getSheet()
    const before = sheet?.images.length ?? 0
    this.ensureBaseline()
    const inv = this.exec({ kind: 'convert-to-cells', targets: allowed }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    const after = this.host.getSheet()?.images.length ?? before
    return Math.max(0, before - after)
  }
  /**
   * 单元格内嵌图 → 浮动图:把 (row,col) 的 DISPIMG 拎出来变成浮动图(默认 96×96px)。
   * 非内嵌图格 → 返 false。入命令栈。
   * Phase A 补漏: 源格只读 → 拒绝 (拎出来意味着该格内容变了,属于修改).
   */
  convertCellImageToFloat(row: number, col: number, size?: { width: number; height: number }): boolean {
    if (!this.host.isEditingEnabled()) return false
    if (!this.host.isEditable(row, col)) {
      this.host.emit('permission-denied', { reason: 'image-convert', cells: [{ row, col }], message: '源格只读,无法转浮动图' } satisfies PermissionDeniedPayload)
      return false
    }
    this.ensureBaseline()
    const inv = this.exec({ kind: 'convert-to-float', row, col, size }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    return !!inv
  }
  /**
   * 批量把多格 DISPIMG 拎成浮动图(选区批量反向);非内嵌图格自动跳过;一次进撤销栈。
   * 返回成功转换的张数。
   * Phase A 补漏: 过滤掉源格只读的, denied 一次性 emit.
   */
  convertCellImagesToFloats(cells: { row: number; col: number; size?: { width: number; height: number } }[]): number {
    if (!this.host.isEditingEnabled() || !cells.length) return 0
    const denied: Array<{ row: number; col: number }> = []
    const allowed: typeof cells = []
    for (const c of cells) {
      if (this.host.isEditable(c.row, c.col)) allowed.push(c)
      else denied.push({ row: c.row, col: c.col })
    }
    if (denied.length) this.host.emit('permission-denied', { reason: 'image-convert', cells: denied, message: `${denied.length} 个源格只读,已跳过` } satisfies PermissionDeniedPayload)
    if (!allowed.length) return 0
    const sheet = this.host.getSheet()
    const before = sheet?.images.length ?? 0
    this.ensureBaseline()
    const inv = this.exec({ kind: 'convert-to-floats', cells: allowed }, 'api')
    if (inv) {
      this.pushUndo(inv)
      this.markDirty()
    }
    const after = this.host.getSheet()?.images.length ?? before
    return Math.max(0, after - before)
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
   * 批量设维度尺寸 (Phase B, 2026-06-08). target 接 number | number[] | {from,to}.
   * 多 index 时聚合成单次 restore-wb undo (跟 convertImagesToCells 范式一致).
   * strictDimensions=true + 白名单未覆盖某 index → skip 那个 + 累计 deniedDims; 整次结束 emit 一次 permission-denied.
   * 返回成功条数 (0 = 全 skip / editable=false).
   */
  setDimensions(axis: DimAxis, indices: number[], size: number, canDim: (i: number) => boolean): number {
    if (!this.host.isEditingEnabled() || !indices.length) return 0
    const sheet = this.host.getSheet()
    const wb = this.host.getWorkbook()
    if (!sheet || !wb) return 0
    const allowed: number[] = []
    const denied: number[] = []
    for (const i of indices) {
      if (i < 0) continue
      if (canDim(i)) allowed.push(i)
      else denied.push(i)
    }
    if (denied.length) {
      this.host.emit('permission-denied', { reason: 'dimension', cells: [], dims: { axis: axis as 'col' | 'row', indices: denied }, message: `${denied.length} 个 ${axis} 未覆盖白名单,已跳过` } satisfies PermissionDeniedPayload)
    }
    if (!allowed.length) return 0
    // 单条直接走 setDimension (复用现有 dim-change emit + 命令栈)
    if (allowed.length === 1) {
      return this.setDimension(axis, allowed[0], size) ? 1 : 0
    }
    // 多条: 一次 snap + 循环 setCellSize-via-mutation + 单 restore-wb undo
    this.ensureBaseline()
    const snap = cloneWorkbook(wb)
    let written = 0
    for (const i of allowed) {
      const beforeSize = currentDimSize(sheet, axis, i)
      // 直接走 mutation, 不入单条命令栈
      if (axis === 'col') {
        const info = sheet.columns.get(i)
        sheet.columns.set(i, { width: size, hidden: info?.hidden ?? false })
      } else {
        const info = sheet.rows.get(i)
        sheet.rows.set(i, { height: size, hidden: info?.hidden ?? false })
      }
      written++
      this.host.emit('dim-change', { axis, index: i, before: beforeSize, after: size, source: 'api' } satisfies DimChangePayload)
    }
    this.pushUndo({ kind: 'restore-wb', snapshot: snap })
    this.markDirty()
    this.host.onModelChange()
    return written
  }

  /**
   * 重置某些列宽/行高到默认 (Phase B, 2026-06-08) — 移除 columns/rows Map 条目, 渲染回落 defaultColWidth/Height.
   * 多 index 时单次 restore-wb undo.
   */
  resetDimensions(axis: DimAxis, indices: number[], canDim: (i: number) => boolean): number {
    if (!this.host.isEditingEnabled() || !indices.length) return 0
    const sheet = this.host.getSheet()
    const wb = this.host.getWorkbook()
    if (!sheet || !wb) return 0
    const allowed: number[] = []
    const denied: number[] = []
    for (const i of indices) {
      if (i < 0) continue
      if (canDim(i)) allowed.push(i)
      else denied.push(i)
    }
    if (denied.length) {
      this.host.emit('permission-denied', { reason: 'dimension', cells: [], dims: { axis: axis as 'col' | 'row', indices: denied }, message: `${denied.length} 个 ${axis} 未覆盖白名单,已跳过 reset` } satisfies PermissionDeniedPayload)
    }
    if (!allowed.length) return 0
    this.ensureBaseline()
    const snap = cloneWorkbook(wb)
    let written = 0
    const defaultSize = axis === 'col' ? sheet.defaultColWidth : sheet.defaultRowHeight
    for (const i of allowed) {
      const beforeSize = currentDimSize(sheet, axis, i)
      if (axis === 'col') sheet.columns.delete(i)
      else sheet.rows.delete(i)
      written++
      if (beforeSize !== defaultSize) {
        this.host.emit('dim-change', { axis, index: i, before: beforeSize, after: defaultSize, source: 'api' } satisfies DimChangePayload)
      }
    }
    this.pushUndo({ kind: 'restore-wb', snapshot: snap })
    this.markDirty()
    this.host.onModelChange()
    return written
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
  /** 公开版 pushUndo (Phase B 2026-06-08): 给 controller 的批量 autoFit 用. 一般业务不要直调. */
  pushUndoExternal(inverse: EditCommand): void {
    this.pushUndo(inverse)
  }
  /** 公开版 markDirty: 给 controller 的批量 autoFit 用. 一般业务不要直调. */
  markDirtyExternal(): void {
    this.markDirty()
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
    // 转换族(WPS 内嵌图 ⇄ 浮动图):整簿快照逆(动 cells + images + cellImages 登记表)
    if (cmd.kind === 'convert-to-cell' || cmd.kind === 'convert-to-cells' || cmd.kind === 'convert-to-float' || cmd.kind === 'convert-to-floats') {
      const wb = this.host.getWorkbook()
      if (!wb) return null
      const d = this.host.getDate1904()
      // 受影响格(单/批/拆/批拆):先拍前态快照,apply 后逐格拍后态发 cell-change
      const cells: CellPos[] =
        cmd.kind === 'convert-to-cells' ? cmd.targets.map((t) => ({ row: t.row, col: t.col }))
        : cmd.kind === 'convert-to-floats' ? cmd.cells.map((c) => ({ row: c.row, col: c.col }))
        : [{ row: cmd.row, col: cmd.col }]
      const before = cells.map((p) => buildCellSnapshot(sheet, p.row, p.col, d))
      const snap = cloneWorkbook(wb)
      let any = false
      if (cmd.kind === 'convert-to-cell') {
        any = convertFloatToCellImage(wb, sheet, cmd.imageIndex, cmd.row, cmd.col) !== null
      } else if (cmd.kind === 'convert-to-float') {
        any = convertCellImageToFloat(wb, sheet, cmd.row, cmd.col, cmd.size) !== -1
      } else if (cmd.kind === 'convert-to-floats') {
        // 反向批量(嵌入→浮动):逐格调,聚合成单次撤销
        for (const c of cmd.cells) if (convertCellImageToFloat(wb, sheet, c.row, c.col, c.size) !== -1) any = true
      } else {
        // 批量:按 imageIndex 降序处理(先删高索引,低索引不受 splice 影响)
        const sorted = [...cmd.targets].sort((a, b) => b.imageIndex - a.imageIndex)
        for (const t of sorted) if (convertFloatToCellImage(wb, sheet, t.imageIndex, t.row, t.col) !== null) any = true
      }
      if (!any) return null // 一张都没转(缺字节 / 非内嵌图格)→ 不入栈、不动模型
      this.host.rebuildOverlays() // 先重建叠加层(浮动图增删 → imageEls 与 sheet.images 重新对齐)
      this.host.onModelChange() // 再重建几何 + 重绘(此时 position 读到的是新 imageEls)
      if (this.host.isRecalcEnabled()) this.refreshEngine() // 公式集变了,重建引擎
      cells.forEach((p, i) =>
        this.host.emit('cell-change', { before: before[i], after: buildCellSnapshot(sheet, p.row, p.col, d), source } satisfies CellChangePayload),
      )
      this.host.emit('image-change', { index: -1, before: null, after: null, source } satisfies ImageChangePayload)
      return { kind: 'restore-wb', snapshot: snap }
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
