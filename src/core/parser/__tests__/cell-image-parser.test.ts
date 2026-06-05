import { describe, it, expect } from 'vitest'
import { attachCellImages, dispImgIdOf } from '../cell-image-parser'
import type { RawPackage } from '../raw-xml'
import { makeSheet } from '../../__tests__/helpers'
import type { CellModel, SheetModel } from '../../model/types'
import { cellKey } from '../../model/types'

/** 构造一个最小 RawPackage,只回 cellimages.xml / 其 rels / media 字节 */
function fakePkg(opts: { withRegistry?: boolean } = {}): RawPackage {
  const cellImagesXml = {
    cellImages: {
      cellImage: [
        { pic: { nvPicPr: { cNvPr: { '@_name': 'ID_aaa' } }, blipFill: { blip: { '@_embed': 'rId1' } } } },
        { pic: { nvPicPr: { cNvPr: { '@_name': 'ID_bbb' } }, blipFill: { blip: { '@_embed': 'rId2' } } } },
      ],
    },
  }
  const relsXml = {
    Relationships: {
      Relationship: [
        { '@_Id': 'rId1', '@_Target': 'media/image1.png' },
        { '@_Id': 'rId2', '@_Target': 'media/image2.jpeg' },
      ],
    },
  }
  const files: Record<string, Uint8Array> = {
    'xl/media/image1.png': new Uint8Array([1, 2, 3]),
    'xl/media/image2.jpeg': new Uint8Array([4, 5, 6, 7]),
  }
  return {
    files,
    bytes: (p) => files[p.replace(/^\//, '')],
    text: () => undefined,
    parse: (p) => {
      const k = p.replace(/^\//, '')
      if (!opts.withRegistry) return undefined
      if (k === 'xl/cellimages.xml') return cellImagesXml
      if (k === 'xl/_rels/cellimages.xml.rels') return relsXml
      return undefined
    },
    list: () => [],
  }
}

function dispCell(row: number, col: number, formula: string): CellModel {
  return { row, col, type: 'formula', raw: null, formula, styleId: 0 }
}

describe('dispImgIdOf', () => {
  it('抽取 WPS 各种写法的 DISPIMG id', () => {
    expect(dispImgIdOf('_xlfn.DISPIMG("ID_aaa",1)')).toBe('ID_aaa')
    expect(dispImgIdOf('DISPIMG("ID_bbb", 1)')).toBe('ID_bbb')
    expect(dispImgIdOf('=DISPIMG("ID_ccc",1)')).toBe('ID_ccc')
  })
  it('非 DISPIMG 公式返 undefined', () => {
    expect(dispImgIdOf('A1+A2')).toBeUndefined()
    expect(dispImgIdOf('SUM(A1:A5)')).toBeUndefined()
    expect(dispImgIdOf(undefined)).toBeUndefined()
  })
})

describe('attachCellImages', () => {
  it('建登记表 + 回填单元格 dispImgId(bytes/mime 正确)', () => {
    const sheet = makeSheet()
    sheet.cells.set(cellKey(0, 0), dispCell(0, 0, '_xlfn.DISPIMG("ID_aaa",1)'))
    sheet.cells.set(cellKey(1, 0), dispCell(1, 0, 'DISPIMG("ID_bbb",1)'))
    sheet.cells.set(cellKey(2, 0), dispCell(2, 0, 'SUM(A1:A2)')) // 普通公式,不动

    const reg = attachCellImages(fakePkg({ withRegistry: true }), [sheet])

    expect(reg).toBeDefined()
    expect(reg!.size).toBe(2)
    expect(reg!.get('ID_aaa')).toMatchObject({ id: 'ID_aaa', mime: 'image/png' })
    expect(Array.from(reg!.get('ID_aaa')!.bytes!)).toEqual([1, 2, 3])
    expect(reg!.get('ID_bbb')!.mime).toBe('image/jpeg')

    expect(sheet.cells.get(cellKey(0, 0))!.dispImgId).toBe('ID_aaa')
    expect(sheet.cells.get(cellKey(1, 0))!.dispImgId).toBe('ID_bbb')
    expect(sheet.cells.get(cellKey(2, 0))!.dispImgId).toBeUndefined()
  })

  it('无 cellimages.xml(非 WPS 文件)→ 返 undefined,不标记任何格', () => {
    const sheet: SheetModel = makeSheet()
    sheet.cells.set(cellKey(0, 0), dispCell(0, 0, 'A1+1'))
    const reg = attachCellImages(fakePkg({ withRegistry: false }), [sheet])
    expect(reg).toBeUndefined()
    expect(sheet.cells.get(cellKey(0, 0))!.dispImgId).toBeUndefined()
  })

  it('有 DISPIMG 公式但无登记表 → 仍标 dispImgId(渲染层画占位),返空表', () => {
    const sheet = makeSheet()
    sheet.cells.set(cellKey(0, 0), dispCell(0, 0, 'DISPIMG("ID_zzz",1)'))
    const reg = attachCellImages(fakePkg({ withRegistry: false }), [sheet])
    expect(reg).toBeDefined()
    expect(reg!.size).toBe(0)
    expect(sheet.cells.get(cellKey(0, 0))!.dispImgId).toBe('ID_zzz')
  })
})
