import { describe, it, expect } from 'vitest'
import {
  insertRows,
  deleteRows,
  insertCols,
  deleteCols,
  captureStructure,
  restoreStructure,
  deleteIntersectsMerge,
} from '../structure'
import type { CellModel, ImageAnchor, SheetModel } from '../types'
import { cellKey } from '../types'

function cell(row: number, col: number, raw: CellModel['raw']): CellModel {
  return { row, col, type: 'number', raw, styleId: 0 }
}
function sheet(): SheetModel {
  const cells = new Map<string, CellModel>()
  const put = (r: number, c: number, v: number) => cells.set(cellKey(r, c), cell(r, c, v))
  put(0, 0, 1)
  put(2, 0, 3) // 第 2 行
  put(2, 1, 30)
  const img: ImageAnchor = { src: 'x', from: { col: 2, colOffEmu: 0, row: 2, rowOffEmu: 0 } }
  return {
    cells,
    styles: [],
    merges: [{ top: 2, left: 0, bottom: 2, right: 1 }], // 第 2 行的 A2:B2 合并
    rows: new Map([[2, { height: 30, hidden: false }]]),
    columns: new Map([[1, { width: 80, hidden: false }]]),
    images: [img],
    dimension: { rows: 3, cols: 2 },
  } as unknown as SheetModel
}

describe('structure 插入行(E7)', () => {
  it('insertRows:>=at 的格/合并/行高/图片下移,dimension+1', () => {
    const s = sheet()
    insertRows(s, 1, 1) // 在第 1 行插入 1 行
    expect(s.cells.get(cellKey(0, 0))).toMatchObject({ raw: 1 }) // 第 0 行不动
    expect(s.cells.has(cellKey(2, 0))).toBe(false) // 原第 2 行已移走
    expect(s.cells.get(cellKey(3, 0))).toMatchObject({ row: 3, raw: 3 }) // 下移到第 3 行
    expect(s.merges[0]).toMatchObject({ top: 3, bottom: 3 }) // 合并下移
    expect(s.rows.get(3)).toMatchObject({ height: 30 }) // 行高随之
    expect(s.images[0].from.row).toBe(3) // 图片锚点下移
    expect(s.dimension.rows).toBe(4)
  })

  it('insertRows:跨插入点的合并区 → 扩展(只 bottom 下移)', () => {
    const s = sheet()
    s.merges = [{ top: 1, left: 0, bottom: 2, right: 0 }] // 跨第 1 行
    insertRows(s, 2, 1)
    expect(s.merges[0]).toMatchObject({ top: 1, bottom: 3 }) // top 不动,bottom 扩展
  })
})

describe('structure 删除行(E7)', () => {
  it('deleteRows:删除段格移除、之后上移,相交合并丢弃,dimension-1', () => {
    const s = sheet()
    deleteRows(s, 2, 1) // 删第 2 行(含合并 + 数据)
    expect(s.cells.has(cellKey(2, 0))).toBe(false) // 第 2 行数据没了
    expect(s.merges).toHaveLength(0) // 相交合并被丢弃
    expect(s.rows.has(2)).toBe(false)
    expect(s.dimension.rows).toBe(2)
  })

  it('deleteRows:删除段之上的行不动,之下上移', () => {
    const s = sheet()
    s.cells.set(cellKey(4, 0), cell(4, 0, 5))
    s.dimension.rows = 5
    deleteRows(s, 1, 1) // 删第 1 行(空)
    expect(s.cells.get(cellKey(0, 0))).toMatchObject({ raw: 1 }) // 第 0 行不动
    expect(s.cells.get(cellKey(1, 0))).toMatchObject({ raw: 3 }) // 原第 2 行上移到 1
    expect(s.cells.get(cellKey(3, 0))).toMatchObject({ raw: 5 }) // 原第 4 行上移到 3
  })
})

describe('structure 插入/删除列(E7)', () => {
  it('insertCols / deleteCols 对称移位', () => {
    const s = sheet()
    insertCols(s, 1, 2) // 第 1 列前插 2 列
    expect(s.cells.get(cellKey(2, 3))).toMatchObject({ raw: 30 }) // 原 (2,1) → (2,3)
    expect(s.columns.get(3)).toMatchObject({ width: 80 }) // 列宽随之
    expect(s.dimension.cols).toBe(4)
    deleteCols(s, 1, 2) // 还原
    expect(s.cells.get(cellKey(2, 1))).toMatchObject({ raw: 30 })
    expect(s.dimension.cols).toBe(2)
  })
})

describe('structure 快照逆 + 守卫', () => {
  it('captureStructure/restoreStructure:删除后还原 = 原状(含合并/行高/图片)', () => {
    const s = sheet()
    const snap = captureStructure(s)
    deleteRows(s, 2, 1)
    restoreStructure(s, snap)
    expect(s.cells.get(cellKey(2, 0))).toMatchObject({ raw: 3 }) // 数据回来
    expect(s.merges).toHaveLength(1) // 合并回来
    expect(s.rows.get(2)).toMatchObject({ height: 30 })
    expect(s.images[0].from.row).toBe(2)
    expect(s.dimension.rows).toBe(3)
  })

  it('deleteIntersectsMerge:删除段跨越合并区时为真', () => {
    const s = sheet() // 合并 A2:B2(row 2)
    expect(deleteIntersectsMerge(s, 'delete-rows', 2, 1)).toBe(false) // 正好整行删 = 不算"跨越"(整个被删)
    s.merges = [{ top: 1, left: 0, bottom: 3, right: 0 }]
    expect(deleteIntersectsMerge(s, 'delete-rows', 2, 1)).toBe(true) // 删中间一行 → 跨越
  })
})
