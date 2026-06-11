import { describe, it, expect } from 'vitest'
import { attachPivotTables } from '../pivot-parser'
import type { RawPackage } from '../raw-xml'
import { makeSheet } from '../../__tests__/helpers'

function fakePkg(): RawPackage {
  const parsed: Record<string, unknown> = {
    'xl/workbook.xml': {
      workbook: {
        sheets: { sheet: { '@_name': 'Pivot', '@_id': 'rId1' } },
        pivotCaches: { pivotCache: { '@_cacheId': 3, '@_id': 'rIdCache' } },
      },
    },
    'xl/_rels/workbook.xml.rels': {
      Relationships: {
        Relationship: [
          { '@_Id': 'rId1', '@_Target': 'worksheets/sheet1.xml' },
          { '@_Id': 'rIdCache', '@_Target': 'pivotCache/pivotCacheDefinition1.xml' },
        ],
      },
    },
    'xl/worksheets/sheet1.xml': {
      worksheet: { pivotTableDefinition: { '@_id': 'rIdPivot' } },
    },
    'xl/worksheets/_rels/sheet1.xml.rels': {
      Relationships: {
        Relationship: { '@_Id': 'rIdPivot', '@_Target': '../pivotTables/pivotTable1.xml' },
      },
    },
    'xl/pivotTables/pivotTable1.xml': {
      pivotTableDefinition: {
        '@_name': 'PivotTable1',
        '@_cacheId': 3,
        location: { '@_ref': 'C4:F12', '@_firstHeaderRow': 1, '@_firstDataRow': 3, '@_firstDataCol': 2 },
        rowFields: { field: { '@_x': 0 } },
        colFields: { field: { '@_x': 1 } },
        pageFields: { pageField: { '@_fld': 2 } },
        dataFields: { dataField: { '@_fld': 3, '@_name': '求和项: 金额' } },
      },
    },
    'xl/pivotCache/pivotCacheDefinition1.xml': {
      pivotCacheDefinition: {
        cacheFields: {
          cacheField: [
            { '@_name': '地区' },
            { '@_name': '月份' },
            { '@_name': '产品' },
            { '@_name': '金额' },
          ],
        },
      },
    },
  }
  return {
    files: {},
    bytes: () => undefined,
    text: () => undefined,
    parse: (p) => parsed[p.replace(/^\//, '')],
    list: () => [],
  }
}

describe('attachPivotTables', () => {
  it('解析透视表范围、缓存字段和只读字段按钮', () => {
    const sheet = makeSheet({ name: 'Pivot' })

    attachPivotTables(fakePkg(), [sheet])

    expect(sheet.pivotTables).toHaveLength(1)
    const pivot = sheet.pivotTables[0]
    expect(pivot.name).toBe('PivotTable1')
    expect(pivot.range).toEqual({ top: 3, left: 2, bottom: 11, right: 5 })
    expect(pivot.fields).toEqual(['地区', '月份', '产品', '金额'])
    expect(pivot.buttons).toEqual(expect.arrayContaining([
      { row: 3, col: 2, label: '产品', kind: 'page' },
      { row: 4, col: 4, label: '月份', kind: 'col' },
      { row: 5, col: 2, label: '地区', kind: 'row' },
      { row: 5, col: 4, label: '求和项: 金额', kind: 'data' },
    ]))
  })

  it('sheet 名对不上 workbook 时不回填', () => {
    const sheet = makeSheet({ name: 'Other' })
    attachPivotTables(fakePkg(), [sheet])
    expect(sheet.pivotTables).toEqual([])
  })
})
