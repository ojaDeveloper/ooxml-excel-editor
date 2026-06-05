import { describe, it, expect } from 'vitest'
import { EditController, type EditControllerHost } from '../edit-controller'
import type { CellStyle, SheetModel, WorkbookModel } from '../../model/types'
import type { CellValue } from '../../model/data-access'
import type { FormulaEngine } from '../../formula/engine'
import { cellKey } from '../../model/types'

const styleA = { font: {}, fill: { type: 'none' }, numFmt: 'General' } as unknown as CellStyle

function setup(readOnly: Set<string> = new Set(), editingEnabled = true, engine: FormulaEngine | null = null) {
  const sheet = {
    name: 'S1',
    cells: new Map(),
    styles: [styleA],
    columns: new Map(),
    rows: new Map(),
    merges: [],
    images: [],
    defaultColWidth: 64,
    defaultRowHeight: 20,
    dimension: { rows: 0, cols: 0 },
  } as unknown as SheetModel
  const workbook = { sheets: [sheet], activeSheet: 0, themeColors: [], date1904: false } as unknown as WorkbookModel
  const events: { event: string; payload: any }[] = []
  let renders = 0
  const host: EditControllerHost = {
    getSheet: () => sheet,
    getWorkbook: () => workbook,
    getDate1904: () => false,
    isEditable: (r, c) => editingEnabled && !readOnly.has(`${r}:${c}`),
    isEditingEnabled: () => editingEnabled,
    getActiveSheetIndex: () => 0,
    isRecalcEnabled: () => !!engine,
    getEngineFactory: () => (engine ? async () => engine : null),
    onModelChange: () => {
      renders++
    },
    emit: (event, payload) => events.push({ event, payload }),
  }
  return { sheet, workbook, ec: new EditController(host), events, renders: () => renders }
}

describe('EditController(命令式编辑 + 前后快照事件 + undo/redo)', () => {
  it('editCell:改模型 + 发 cell-change(前空→后值)+ 触发重绘', () => {
    const { sheet, ec, events, renders } = setup()
    const ok = ec.editCell(0, 0, 42)
    expect(ok).toBe(true)
    expect(sheet.cells.get(cellKey(0, 0))).toMatchObject({ type: 'number', raw: 42 })
    expect(renders()).toBe(1)
    const cc = events.filter((e) => e.event === 'cell-change')
    expect(cc).toHaveLength(1)
    expect(cc[0].payload.source).toBe('api')
    expect(cc[0].payload.before.cell).toBeNull() // 前态空
    expect(cc[0].payload.after.raw).toBe(42) // 后态值
    expect(cc[0].payload.after.text).toBe('42') // 显示文本一致
  })

  it('只读格:editCell 不生效、无事件', () => {
    const { sheet, ec, events } = setup(new Set(['0:0']))
    expect(ec.editCell(0, 0, 1)).toBe(false)
    expect(sheet.cells.has(cellKey(0, 0))).toBe(false)
    expect(events).toHaveLength(0)
  })

  it('undo/redo:精确还原前态 + 重做;source 标注', () => {
    const { sheet, ec, events } = setup()
    ec.editCell(1, 1, 'hi')
    expect(ec.canUndo()).toBe(true)
    ec.undo()
    expect(sheet.cells.has(cellKey(1, 1))).toBe(false) // 还原为空
    expect(ec.canRedo()).toBe(true)
    const undoEvt = events.find((e) => e.payload.source === 'undo')
    expect(undoEvt?.payload.before.raw).toBe('hi')
    expect(undoEvt?.payload.after.cell).toBeNull()
    ec.redo()
    expect(sheet.cells.get(cellKey(1, 1))).toMatchObject({ raw: 'hi' }) // 重做回来
    expect(events.some((e) => e.payload.source === 'redo')).toBe(true)
  })

  it('改值覆盖:undo 还回旧值(非删除)', () => {
    const { sheet, ec } = setup()
    ec.editCell(0, 0, 1)
    ec.editCell(0, 0, 2)
    ec.undo()
    expect(sheet.cells.get(cellKey(0, 0))).toMatchObject({ raw: 1 }) // 回到 1,不是空
  })

  it('新编辑清空 redo 栈', () => {
    const { ec } = setup()
    ec.editCell(0, 0, 1)
    ec.undo()
    expect(ec.canRedo()).toBe(true)
    ec.editCell(0, 1, 2) // 新编辑
    expect(ec.canRedo()).toBe(false)
  })

  it('editRange:跳过只读格,只改可编辑格', () => {
    const { sheet, ec } = setup(new Set(['0:0'])) // (0,0) 只读
    const ok = ec.editRange({ top: 0, left: 0, bottom: 0, right: 1 }, [[10, 20]])
    expect(ok).toBe(true)
    expect(sheet.cells.has(cellKey(0, 0))).toBe(false) // 只读跳过
    expect(sheet.cells.get(cellKey(0, 1))).toMatchObject({ raw: 20 }) // 可编辑改了
  })

  it('getCellSnapshot 查询:含 raw/computed/text/完整 cell', () => {
    const { ec } = setup()
    ec.editCell(0, 0, 7)
    const snap = ec.getCellSnapshot(0, 0)!
    expect(snap.raw).toBe(7)
    expect(snap.computed).toBe(7)
    expect(snap.text).toBe('7')
    expect(snap.cell).toMatchObject({ type: 'number', raw: 7 })
  })
})

describe('EditController — 维度编辑(resize 入命令栈;E3.5)', () => {
  it('setDimension:写模型 + 发 dim-change(before 默认→after)+ 可撤销', () => {
    const { sheet, ec, events, renders } = setup()
    const ok = ec.setDimension('col', 2, 120)
    expect(ok).toBe(true)
    expect(sheet.columns.get(2)).toMatchObject({ width: 120, hidden: false })
    expect(renders()).toBe(1)
    const dim = events.find((e) => e.event === 'dim-change')!
    expect(dim.payload).toMatchObject({ axis: 'col', index: 2, before: 64, after: 120, source: 'api' })
    // undo → 回落默认(删 Map 项)
    ec.undo()
    expect(sheet.columns.has(2)).toBe(false)
    expect(ec.canRedo()).toBe(true)
    ec.redo()
    expect(sheet.columns.get(2)).toMatchObject({ width: 120 })
  })

  it('recordDimEdit:模型已改(模拟拖拽),补登 undo 还原前态 + 发 dim-change', () => {
    const { sheet, ec, events } = setup()
    sheet.columns.set(1, { width: 200, hidden: false }) // 模拟 renderer 拖拽已改完
    ec.recordDimEdit('col', 1, null, 64, 200) // 前态无项(null)
    expect(events.some((e) => e.event === 'dim-change')).toBe(true)
    expect(ec.canUndo()).toBe(true)
    ec.undo() // 还原到前态 null → 删项回落默认
    expect(sheet.columns.has(1)).toBe(false)
  })

  it('行高同理:setDimension(row) 改 rows Map', () => {
    const { sheet, ec } = setup()
    ec.setDimension('row', 0, 40)
    expect(sheet.rows.get(0)).toMatchObject({ height: 40 })
  })

  it('非编辑模式:setDimension 不生效', () => {
    const { sheet, ec } = setup(new Set(), false)
    expect(ec.setDimension('col', 0, 99)).toBe(false)
    expect(sheet.columns.has(0)).toBe(false)
  })
})

describe('EditController — 脏状态 + 还原原件(E3.5)', () => {
  it('编辑 → dirty=true + 发 dirty-change;首次才发', () => {
    const { ec, events } = setup()
    expect(ec.isDirty()).toBe(false)
    ec.editCell(0, 0, 1)
    expect(ec.isDirty()).toBe(true)
    const dirtyEvts = events.filter((e) => e.event === 'dirty-change')
    expect(dirtyEvts).toHaveLength(1)
    expect(dirtyEvts[0].payload).toEqual({ dirty: true })
    ec.editCell(0, 1, 2) // 第二次编辑不再发 dirty-change
    expect(events.filter((e) => e.event === 'dirty-change')).toHaveLength(1)
  })

  it('resetToOriginal:还原值 + 列宽 + 清脏 + 清命令栈', () => {
    const { sheet, ec, events } = setup()
    ec.editCell(0, 0, 'edited') // 触发懒捕获 baseline(此刻 sheet 为空)
    ec.setDimension('col', 0, 150)
    expect(ec.isDirty()).toBe(true)
    const ok = ec.resetToOriginal()
    expect(ok).toBe(true)
    expect(sheet.cells.has(cellKey(0, 0))).toBe(false) // 值还原(原本空)
    expect(sheet.columns.has(0)).toBe(false) // 列宽还原(原本默认)
    expect(ec.isDirty()).toBe(false)
    expect(ec.canUndo()).toBe(false)
    expect(events.some((e) => e.event === 'dirty-change' && e.payload.dirty === false)).toBe(true)
  })

  it('resetDirtyBaseline:换新簿作废 baseline,reset 不再还原', () => {
    const { ec } = setup()
    ec.editCell(0, 0, 1)
    ec.resetDirtyBaseline() // 模拟换新工作簿
    expect(ec.isDirty()).toBe(false)
    expect(ec.resetToOriginal()).toBe(false) // 无 baseline
  })
})

/** 极简 mock 引擎:模型化 B1(0,1) = A1(0,0) + 1,够测 EditController 重算编排(不依赖 hyperformula)。 */
function mockEngine(): FormulaEngine {
  const k = (s: number, r: number, c: number) => `${s}:${r}:${c}`
  const vals = new Map<string, CellValue>()
  const recomputeB1 = (out: { sheet: number; row: number; col: number; value: CellValue }[]) => {
    const a1 = vals.get(k(0, 0, 0))
    if (typeof a1 === 'number') {
      const b1 = a1 + 1
      vals.set(k(0, 0, 1), b1)
      out.push({ sheet: 0, row: 0, col: 1, value: b1 })
    }
  }
  return {
    setSheets(wb) {
      const a1 = wb.sheets[0].cells.get(cellKey(0, 0))
      vals.set(k(0, 0, 0), a1 ? (a1.raw as CellValue) : null)
    },
    setCell(s, r, c, content) {
      const v = typeof content === 'string' && content.startsWith('=') ? null : (content as CellValue)
      vals.set(k(s, r, c), v)
      const out = [{ sheet: s, row: r, col: c, value: v }]
      recomputeB1(out)
      return out
    },
    getValue: (s, r, c) => vals.get(k(s, r, c)) ?? null,
    destroy() {},
  }
}

describe('EditController — 公式重算级联(E4;mock 引擎)', () => {
  it('编辑被引用格 → 依赖格自动重算 + 发 cell-change;undo 反向重算', async () => {
    const { sheet, ec, events } = setup(new Set(), true, mockEngine())
    // 种子:A1=5(值),B1=公式 =A1+1(缓存 6)
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'number', raw: 5, styleId: 0 } as never)
    sheet.cells.set(cellKey(0, 1), { row: 0, col: 1, type: 'formula', formula: '=A1+1', raw: 6, styleId: 0 } as never)
    await ec.warmEngine()

    // 改 A1 → 10:B1 应级联到 11 + 发 B1 的 cell-change
    ec.editCell(0, 0, 10)
    expect(sheet.cells.get(cellKey(0, 1))!.raw).toBe(11) // 依赖格重算写回
    const b1Evt = events.filter((e) => e.event === 'cell-change' && e.payload.after?.col === 1)
    expect(b1Evt.length).toBeGreaterThanOrEqual(1)
    expect(b1Evt.at(-1)!.payload.before.raw).toBe(6) // 前态(旧缓存)
    expect(b1Evt.at(-1)!.payload.after.raw).toBe(11) // 后态(重算)

    // undo → A1 回 5,B1 反向重算回 6
    ec.undo()
    expect(sheet.cells.get(cellKey(0, 0))!.raw).toBe(5)
    expect(sheet.cells.get(cellKey(0, 1))!.raw).toBe(6)
  })

  it('引擎未就绪(未 warm)→ 编辑不重算(降级,不报错)', () => {
    const { sheet, ec } = setup(new Set(), true, mockEngine())
    sheet.cells.set(cellKey(0, 1), { row: 0, col: 1, type: 'formula', formula: '=A1+1', raw: 6, styleId: 0 } as never)
    ec.editCell(0, 0, 10) // 未 warmEngine → 引擎未就绪 → B1 不变
    expect(sheet.cells.get(cellKey(0, 1))!.raw).toBe(6)
  })
})
