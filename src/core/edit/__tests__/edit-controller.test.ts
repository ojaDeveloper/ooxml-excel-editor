import { describe, it, expect } from 'vitest'
import { EditController, type EditControllerHost, type CellChangePayload } from '../edit-controller'
import type { CellStyle, SheetModel } from '../../model/types'
import { cellKey } from '../../model/types'

const styleA = { font: {}, fill: { type: 'none' }, numFmt: 'General' } as unknown as CellStyle

function setup(readOnly: Set<string> = new Set()) {
  const sheet = { cells: new Map(), styles: [styleA], dimension: { rows: 0, cols: 0 } } as unknown as SheetModel
  const events: { event: string; payload: CellChangePayload }[] = []
  let renders = 0
  const host: EditControllerHost = {
    getSheet: () => sheet,
    getDate1904: () => false,
    isEditable: (r, c) => !readOnly.has(`${r}:${c}`),
    onModelChange: () => {
      renders++
    },
    emit: (event, payload) => events.push({ event, payload: payload as CellChangePayload }),
  }
  return { sheet, ec: new EditController(host), events, renders: () => renders }
}

describe('EditController(命令式编辑 + 前后快照事件 + undo/redo)', () => {
  it('editCell:改模型 + 发 cell-change(前空→后值)+ 触发重绘', () => {
    const { sheet, ec, events, renders } = setup()
    const ok = ec.editCell(0, 0, 42)
    expect(ok).toBe(true)
    expect(sheet.cells.get(cellKey(0, 0))).toMatchObject({ type: 'number', raw: 42 })
    expect(renders()).toBe(1)
    expect(events).toHaveLength(1)
    expect(events[0].event).toBe('cell-change')
    expect(events[0].payload.source).toBe('api')
    expect(events[0].payload.before.cell).toBeNull() // 前态空
    expect(events[0].payload.after.raw).toBe(42) // 后态值
    expect(events[0].payload.after.text).toBe('42') // 显示文本一致
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
