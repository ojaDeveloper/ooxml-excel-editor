import { describe, it, expect } from 'vitest'
import { inspectCell, findMergeAt, imageAnchorContains } from '../inspect'
import { makeSheet, putNumbers } from '../../__tests__/helpers'
import type { CellModel, ConditionalRule, ImageAnchor, WorkbookModel } from '../types'
import { cellKey } from '../types'

function wb(overlay: Partial<WorkbookModel> = {}): WorkbookModel {
  return {
    sheets: [],
    activeSheet: 0,
    themeColors: Array(17).fill('#000'),
    date1904: false,
    ...overlay,
  } as WorkbookModel
}

describe('Cell Inspector', () => {
  it('普通格:大部分字段为空,snapshot 字段从 buildCellSnapshot 透传', () => {
    const sheet = makeSheet()
    putNumbers(sheet, [{ row: 0, col: 0, v: 42 }])
    const insp = inspectCell(sheet, wb({ sheets: [sheet] }), 0, 0, false)
    expect(insp.row).toBe(0)
    expect(insp.col).toBe(0)
    expect(insp.raw).toBe(42)
    expect(insp.merge).toBeNull()
    expect(insp.isMergeAnchor).toBe(false)
    expect(insp.floatingImages).toEqual([])
    expect(insp.cellImage).toBeNull()
    expect(insp.dataValidation).toBeNull()
    expect(insp.conditional).toEqual([])
    expect(insp.hyperlink).toBeNull()
    expect(insp.comment).toBeNull()
  })

  it('合并区:锚点格 isMergeAnchor=true,被覆盖格 isMergeAnchor=false', () => {
    const sheet = makeSheet({ merges: [{ top: 1, left: 1, bottom: 2, right: 3 }] })
    const wbModel = wb({ sheets: [sheet] })
    const anchor = inspectCell(sheet, wbModel, 1, 1, false)
    expect(anchor.merge).toEqual({ top: 1, left: 1, bottom: 2, right: 3 })
    expect(anchor.isMergeAnchor).toBe(true)

    const covered = inspectCell(sheet, wbModel, 2, 3, false)
    expect(covered.merge).toEqual({ top: 1, left: 1, bottom: 2, right: 3 })
    expect(covered.isMergeAnchor).toBe(false)

    const outside = inspectCell(sheet, wbModel, 0, 0, false)
    expect(outside.merge).toBeNull()
  })

  it('浮动图覆盖:twoCellAnchor [from..to] 内所有格命中,外部不命中', () => {
    const anchorImg: ImageAnchor = {
      src: 'blob:x',
      from: { row: 1, col: 1, colOffEmu: 0, rowOffEmu: 0 },
      to: { row: 3, col: 4, colOffEmu: 0, rowOffEmu: 0 },
    }
    const sheet = makeSheet({ images: [anchorImg] })
    const wbModel = wb({ sheets: [sheet] })
    expect(inspectCell(sheet, wbModel, 2, 2, false).floatingImages.length).toBe(1)
    expect(inspectCell(sheet, wbModel, 1, 1, false).floatingImages.length).toBe(1)
    expect(inspectCell(sheet, wbModel, 3, 4, false).floatingImages.length).toBe(1)
    expect(inspectCell(sheet, wbModel, 0, 0, false).floatingImages.length).toBe(0)
    expect(inspectCell(sheet, wbModel, 4, 4, false).floatingImages.length).toBe(0)
  })

  it('WPS 内嵌图 dispImgId 解出 cellImage', () => {
    const sheet = makeSheet()
    const cell: CellModel = { row: 0, col: 0, type: 'string', raw: '', styleId: 0, dispImgId: 'img1' }
    sheet.cells.set(cellKey(0, 0), cell)
    const wbModel = wb({
      sheets: [sheet],
      cellImages: new Map([['img1', { src: 'blob:abc', mime: 'image/png' }]]),
    } as Partial<WorkbookModel>)
    const insp = inspectCell(sheet, wbModel, 0, 0, false)
    expect(insp.cellImage).toEqual({ id: 'img1', src: 'blob:abc', mime: 'image/png' })

    // 无 dispImgId 的格 → null
    expect(inspectCell(sheet, wbModel, 1, 1, false).cellImage).toBeNull()
  })

  it('数据验证范围命中:返回原 MergeRange 克隆;外部 null', () => {
    const sheet = makeSheet({ dataValidations: [{ top: 0, left: 0, bottom: 5, right: 0 }] })
    const wbModel = wb({ sheets: [sheet] })
    expect(inspectCell(sheet, wbModel, 3, 0, false).dataValidation).toEqual({ top: 0, left: 0, bottom: 5, right: 0 })
    expect(inspectCell(sheet, wbModel, 0, 1, false).dataValidation).toBeNull()
  })

  it('条件格式 cellIs > 100 → 命中规则索引 0,等效样式有 fillColor', () => {
    const rule: ConditionalRule = {
      type: 'cellIs',
      operator: 'greaterThan',
      ranges: [{ top: 0, left: 0, bottom: 9, right: 0 }],
      formulae: ['100'],
      style: { fill: { fgColor: '#FF0000' } },
      priority: 1,
    } as unknown as ConditionalRule
    const sheet = makeSheet({ conditional: [rule] })
    putNumbers(sheet, [{ row: 0, col: 0, v: 150 }, { row: 1, col: 0, v: 50 }])
    const wbModel = wb({ sheets: [sheet] })
    const hit = inspectCell(sheet, wbModel, 0, 0, false)
    expect(hit.conditional).toHaveLength(1)
    expect(hit.conditional[0].ruleIndex).toBe(0)
    expect(hit.conditional[0].style.fill?.fgColor).toBe('#FF0000')
    const miss = inspectCell(sheet, wbModel, 1, 0, false)
    expect(miss.conditional).toEqual([])
  })

  it('hyperlink + comment 透传', () => {
    const sheet = makeSheet()
    const cell: CellModel = {
      row: 0, col: 0, type: 'string', raw: 'link', styleId: 0,
      hyperlink: 'https://example.com',
      comment: 'hello',
    }
    sheet.cells.set(cellKey(0, 0), cell)
    const insp = inspectCell(sheet, wb({ sheets: [sheet] }), 0, 0, false)
    expect(insp.hyperlink).toBe('https://example.com')
    expect(insp.comment).toBe('hello')
  })
})

describe('findMergeAt', () => {
  it('遍历 sheet.merges 找 (r,c) 命中', () => {
    const sheet = makeSheet({ merges: [{ top: 0, left: 0, bottom: 1, right: 1 }, { top: 5, left: 5, bottom: 5, right: 6 }] })
    expect(findMergeAt(sheet, 0, 1)).toEqual({ top: 0, left: 0, bottom: 1, right: 1 })
    expect(findMergeAt(sheet, 5, 6)).toEqual({ top: 5, left: 5, bottom: 5, right: 6 })
    expect(findMergeAt(sheet, 3, 3)).toBeNull()
  })
})

describe('imageAnchorContains', () => {
  it('twoCellAnchor 用矩形测试;oneCellAnchor 用单格测试', () => {
    const two: ImageAnchor = { src: '', from: { row: 1, col: 1, colOffEmu: 0, rowOffEmu: 0 }, to: { row: 3, col: 3, colOffEmu: 0, rowOffEmu: 0 } }
    expect(imageAnchorContains(two, 2, 2)).toBe(true)
    expect(imageAnchorContains(two, 4, 4)).toBe(false)
    const one: ImageAnchor = { src: '', from: { row: 1, col: 1, colOffEmu: 0, rowOffEmu: 0 } }
    expect(imageAnchorContains(one, 1, 1)).toBe(true)
    expect(imageAnchorContains(one, 1, 2)).toBe(false)
  })
})
