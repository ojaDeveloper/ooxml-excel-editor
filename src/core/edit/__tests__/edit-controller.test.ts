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
    images: [] as any[],
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
    rebuildOverlays: () => {
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

  it('setStyle:改 styleId + 发 cell-change(前后 style.font.bold 不同);undo 还原', () => {
    const { sheet, ec, events } = setup()
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'x', styleId: 0 } as never)
    const ok = ec.setStyle({ top: 0, left: 0, bottom: 0, right: 0 }, { font: { bold: true } })
    expect(ok).toBe(true)
    const cc = events.filter((e) => e.event === 'cell-change')
    const last = cc.at(-1)!
    expect(last.payload.before.style.font.bold).toBeFalsy() // 前态非粗
    expect(last.payload.after.style.font.bold).toBe(true) // 后态粗体
    expect(sheet.cells.get(cellKey(0, 0))!.styleId).not.toBe(0)
    ec.undo()
    expect(sheet.cells.get(cellKey(0, 0))!.styleId).toBe(0) // styleId 还原
  })

  it('mergeCells:清空被覆盖格、加合并;undo 还原格与合并(G1)', () => {
    const { sheet, ec } = setup()
    for (const [r, c, v] of [[0, 0, 'A'], [0, 1, 'B'], [1, 0, 'C'], [1, 1, 'D']] as const)
      sheet.cells.set(cellKey(r, c), { row: r, col: c, type: 'string', raw: v, styleId: 0 } as never)
    expect(ec.mergeCells({ top: 0, left: 0, bottom: 1, right: 1 })).toBe(true)
    expect(sheet.merges).toHaveLength(1)
    expect(sheet.cells.get(cellKey(0, 0))!.raw).toBe('A') // 锚点保留
    expect(sheet.cells.has(cellKey(0, 1))).toBe(false) // 被覆盖格清空
    expect(sheet.cells.has(cellKey(1, 1))).toBe(false)
    ec.undo()
    expect(sheet.merges).toHaveLength(0) // 合并撤销
    expect(sheet.cells.get(cellKey(1, 1))!.raw).toBe('D') // 被清格还原
  })

  it('mergeCells:吸收相交旧合并 + 单格不合并', () => {
    const { sheet, ec } = setup()
    sheet.merges.push({ top: 0, left: 0, bottom: 0, right: 1 })
    expect(ec.mergeCells({ top: 0, left: 0, bottom: 1, right: 1 })).toBe(true)
    expect(sheet.merges).toHaveLength(1) // 旧的被吸收,只剩新的
    expect(sheet.merges[0]).toMatchObject({ bottom: 1, right: 1 })
    expect(ec.mergeCells({ top: 5, left: 5, bottom: 5, right: 5 })).toBe(false) // 单格不合并
  })

  it('unmergeCells:移除相交合并;undo 还原', () => {
    const { sheet, ec } = setup()
    sheet.merges.push({ top: 2, left: 0, bottom: 2, right: 3 })
    expect(ec.unmergeCells({ top: 2, left: 1, bottom: 2, right: 1 })).toBe(true)
    expect(sheet.merges).toHaveLength(0)
    ec.undo()
    expect(sheet.merges).toHaveLength(1) // 还原
  })

  it('setStyle:跳过只读格', () => {
    const { sheet, ec } = setup(new Set(['0:0']))
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'x', styleId: 0 } as never)
    expect(ec.setStyle({ top: 0, left: 0, bottom: 0, right: 0 }, { font: { bold: true } })).toBe(false)
    expect(sheet.cells.get(cellKey(0, 0))!.styleId).toBe(0)
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

describe('EditController — 图片编辑(E6)', () => {
  const anchor = () => ({ src: 'data:x', from: { col: 1, colOffEmu: 0, row: 1, rowOffEmu: 0 }, extWidthEmu: 100, extHeightEmu: 100 }) as any

  it('addImage:加图 + 发 image-change(before null→after)+ undo 删除', () => {
    const { sheet, ec, events } = setup()
    const idx = ec.addImage(anchor())
    expect(idx).toBe(0)
    expect(sheet.images).toHaveLength(1)
    const ev = events.find((e) => e.event === 'image-change')!
    expect(ev.payload.before).toBeNull()
    expect(ev.payload.after.src).toBe('data:x')
    expect(ec.isDirty()).toBe(true)
    ec.undo()
    expect(sheet.images).toHaveLength(0) // 加图的逆=删图
  })

  it('removeImage:删图 + 发 image-change(before→after null)+ undo 复原', () => {
    const { sheet, ec, events } = setup()
    sheet.images.push(anchor())
    expect(ec.removeImage(0)).toBe(true)
    expect(sheet.images).toHaveLength(0)
    const ev = events.filter((e) => e.event === 'image-change').at(-1)!
    expect(ev.payload.before.src).toBe('data:x')
    expect(ev.payload.after).toBeNull()
    ec.undo()
    expect(sheet.images).toHaveLength(1) // 删图的逆=按原位加回
    expect(sheet.images[0].src).toBe('data:x')
  })

  it('image-set(recordImageEdit):移动后 undo 还原前态锚点', () => {
    const { sheet, ec } = setup()
    sheet.images.push(anchor())
    const before = { ...sheet.images[0], from: { ...sheet.images[0].from } }
    sheet.images[0].from.colOffEmu = 999999 // 模拟拖拽已改完
    ec.recordImageEdit(0, before as any, { ...sheet.images[0], from: { ...sheet.images[0].from } } as any)
    ec.undo()
    expect(sheet.images[0].from.colOffEmu).toBe(0) // 还原前态
  })

  it('getImages:返回克隆(改返回值不影响模型)', () => {
    const { sheet, ec } = setup()
    sheet.images.push(anchor())
    const imgs = ec.getImages()
    imgs[0].from.colOffEmu = 123
    expect(sheet.images[0].from.colOffEmu).toBe(0) // 模型未受影响
  })
})

describe('EditController — WPS 内嵌图 ⇄ 浮动图互转(第二期)', () => {
  const floatWithBytes = () =>
    ({ src: 'blob:x', bytes: new Uint8Array([1, 2, 3]), mime: 'image/png', from: { col: 0, row: 0, colOffEmu: 0, rowOffEmu: 0 } }) as never

  it('浮动图 → 内嵌图:登记表 + dispImgId + 移除浮动图;undo 整簿还原', () => {
    const { sheet, workbook, ec } = setup()
    sheet.images.push(floatWithBytes())
    expect(ec.convertImageToCell(0, 2, 3)).toBe(true)
    expect(sheet.images).toHaveLength(0)
    const cell = sheet.cells.get(cellKey(2, 3))!
    expect(cell.dispImgId).toBeTruthy()
    expect(cell.formula).toContain('DISPIMG')
    expect(workbook.cellImages!.size).toBe(1)
    ec.undo()
    expect(sheet.images).toHaveLength(1)
    expect(sheet.cells.get(cellKey(2, 3))).toBeUndefined()
    expect(workbook.cellImages?.size ?? 0).toBe(0)
  })

  it('内嵌图 → 浮动图:清空格 + 新浮动图 + 回收无引用登记项;非内嵌格返 false', () => {
    const { sheet, workbook, ec } = setup()
    sheet.images.push(floatWithBytes())
    ec.convertImageToCell(0, 2, 3) // 先造一个内嵌图格
    expect(ec.convertCellImageToFloat(2, 3)).toBe(true)
    expect(sheet.images).toHaveLength(1)
    expect(sheet.cells.get(cellKey(2, 3))).toBeUndefined()
    expect(workbook.cellImages?.size ?? 0).toBe(0) // 无引用 → 回收
    expect(ec.convertCellImageToFloat(5, 5)).toBe(false) // 非内嵌图格
  })

  it('浮动图缺字节 → 不可转,返 false 不动模型', () => {
    const { sheet, ec } = setup()
    sheet.images.push({ src: '', from: { col: 0, row: 0, colOffEmu: 0, rowOffEmu: 0 } } as never)
    expect(ec.convertImageToCell(0, 1, 1)).toBe(false)
    expect(sheet.images).toHaveLength(1)
  })

  it('批量嵌入 convertImagesToCells:多图一次入栈 + 单次 undo 全还原', () => {
    const { sheet, workbook, ec } = setup()
    sheet.images.push(floatWithBytes(), floatWithBytes(), floatWithBytes())
    const n = ec.convertImagesToCells([
      { imageIndex: 0, row: 1, col: 0 },
      { imageIndex: 1, row: 2, col: 0 },
      { imageIndex: 2, row: 3, col: 0 },
    ])
    expect(n).toBe(3)
    expect(sheet.images).toHaveLength(0)
    expect(workbook.cellImages!.size).toBe(3)
    expect(sheet.cells.get(cellKey(1, 0))?.dispImgId).toBeTruthy()
    expect(sheet.cells.get(cellKey(3, 0))?.dispImgId).toBeTruthy()
    // 单次 undo 把整批撤回
    ec.undo()
    expect(sheet.images).toHaveLength(3)
    expect(workbook.cellImages?.size ?? 0).toBe(0)
    expect(sheet.cells.get(cellKey(1, 0))).toBeUndefined()
  })

  it('批量:缺字节的图被跳过,只数成功的', () => {
    const { sheet, ec } = setup()
    sheet.images.push(floatWithBytes(), { src: '', from: { col: 0, row: 0, colOffEmu: 0, rowOffEmu: 0 } } as never)
    const n = ec.convertImagesToCells([
      { imageIndex: 0, row: 1, col: 0 },
      { imageIndex: 1, row: 2, col: 0 },
    ])
    expect(n).toBe(1) // 第二张缺字节 → 跳过
    expect(sheet.images).toHaveLength(1) // 缺字节那张仍在
  })

  it('互转往返 redo:undo 后 redo 重做转换', () => {
    const { sheet, workbook, ec } = setup()
    sheet.images.push(floatWithBytes())
    ec.convertImageToCell(0, 1, 1)
    ec.undo()
    expect(sheet.images).toHaveLength(1)
    ec.redo()
    expect(sheet.images).toHaveLength(0)
    expect(workbook.cellImages!.size).toBe(1)
    expect(sheet.cells.get(cellKey(1, 1))?.dispImgId).toBeTruthy()
  })
})

describe('EditController — 行列结构编辑(E7)', () => {
  it('insertRows:格下移 + 发 struct-change + undo 还原', () => {
    const { sheet, ec, events } = setup()
    sheet.cells.set(cellKey(2, 0), { row: 2, col: 0, type: 'number', raw: 5, styleId: 0 } as never)
    sheet.dimension.rows = 3
    expect(ec.insertRows(1, 1)).toBe(true)
    expect(sheet.cells.has(cellKey(2, 0))).toBe(false)
    expect(sheet.cells.get(cellKey(3, 0))).toMatchObject({ raw: 5 }) // 下移
    const ev = events.find((e) => e.event === 'struct-change')!
    expect(ev.payload).toMatchObject({ op: 'insert-rows', at: 1, count: 1 })
    expect(ec.isDirty()).toBe(true)
    ec.undo()
    expect(sheet.cells.get(cellKey(2, 0))).toMatchObject({ raw: 5 }) // 快照还原
    expect(sheet.cells.has(cellKey(3, 0))).toBe(false)
  })

  it('deleteRows:删数据 + undo 还原被删内容', () => {
    const { sheet, ec } = setup()
    sheet.cells.set(cellKey(2, 0), { row: 2, col: 0, type: 'string', raw: 'gone', styleId: 0 } as never)
    sheet.dimension.rows = 3
    expect(ec.deleteRows(2, 1)).toBe(true)
    expect(sheet.cells.has(cellKey(2, 0))).toBe(false)
    ec.undo()
    expect(sheet.cells.get(cellKey(2, 0))).toMatchObject({ raw: 'gone' }) // 删除内容被还原
  })

  it('insertCols + undo/redo 往返', () => {
    const { sheet, ec } = setup()
    sheet.cells.set(cellKey(0, 1), { row: 0, col: 1, type: 'number', raw: 9, styleId: 0 } as never)
    sheet.dimension.cols = 2
    ec.insertCols(0, 1)
    expect(sheet.cells.get(cellKey(0, 2))).toMatchObject({ raw: 9 }) // 右移
    ec.undo()
    expect(sheet.cells.get(cellKey(0, 1))).toMatchObject({ raw: 9 })
    ec.redo()
    expect(sheet.cells.get(cellKey(0, 2))).toMatchObject({ raw: 9 }) // 重做再右移
  })

  it('非编辑模式:结构编辑不生效', () => {
    const { sheet, ec } = setup(new Set(), false)
    sheet.dimension.rows = 3
    expect(ec.insertRows(0, 1)).toBe(false)
    expect(sheet.dimension.rows).toBe(3)
  })

  it('插入行:公式引用自动重写(=A5→=A6),undo 还原引用文本(F1)', () => {
    const { sheet, ec } = setup()
    sheet.cells.set(cellKey(4, 1), { row: 4, col: 1, type: 'formula', raw: 0, formula: '=A5+1', styleId: 0 } as never)
    sheet.dimension.rows = 5
    ec.insertRows(2, 1) // 在第 2 行上方插入 → 公式格下移到 (5,1),引用 A5→A6
    const moved = sheet.cells.get(cellKey(5, 1))!
    expect(moved.formula).toBe('=A6+1') // 引用重写
    ec.undo()
    expect(sheet.cells.get(cellKey(4, 1))!.formula).toBe('=A5+1') // 还原
  })

  it('删除被引用行:公式引用 → #REF!', () => {
    const { sheet, ec } = setup()
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'formula', raw: 0, formula: '=A5', styleId: 0 } as never)
    sheet.dimension.rows = 5
    ec.deleteRows(4, 1) // 删第 5 行(被 A5 引用)
    expect(sheet.cells.get(cellKey(0, 0))!.formula).toBe('=#REF!')
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
