import { describe, it, expect } from 'vitest'
import { attachRowMeta } from '../row-meta-parser'
import type { RawPackage } from '../raw-xml'
import { makeSheet } from '../../__tests__/helpers'

// 假 pkg:workbook.xml(名→rId)+ rels(rId→sheet 路径)+ sheetN.xml(行 customHeight)
function fakePkg(): RawPackage {
  const parsed: Record<string, unknown> = {
    'xl/workbook.xml': { workbook: { sheets: { sheet: { '@_name': 'Sheet1', '@_id': 'rId1' } } } },
    'xl/_rels/workbook.xml.rels': {
      Relationships: { Relationship: { '@_Id': 'rId1', '@_Target': 'worksheets/sheet1.xml' } },
    },
    'xl/worksheets/sheet1.xml': {
      worksheet: {
        sheetData: {
          row: [
            { '@_r': 1, '@_ht': 66, '@_customHeight': 1 }, // 手动设高
            { '@_r': 2, '@_ht': 30 }, // 自动行(无 customHeight)
            { '@_r': 3, '@_customHeight': '1' }, // 字符串 "1" 也算
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

describe('attachRowMeta', () => {
  it('把 customHeight 标记回填到已存在的 rows 条目', () => {
    const sheet = makeSheet({ name: 'Sheet1' })
    sheet.rows.set(0, { height: 88, hidden: false }) // row1
    sheet.rows.set(1, { height: 40, hidden: false }) // row2
    sheet.rows.set(2, { height: 20, hidden: false }) // row3

    attachRowMeta(fakePkg(), [sheet])

    expect(sheet.rows.get(0)?.customHeight).toBe(true) // customHeight="1"
    expect(sheet.rows.get(1)?.customHeight).toBeUndefined() // 自动行不标
    expect(sheet.rows.get(2)?.customHeight).toBe(true) // 字符串 "1"
  })

  it('sheet 名对不上 workbook → 不动', () => {
    const sheet = makeSheet({ name: '别的表' })
    sheet.rows.set(0, { height: 88, hidden: false })
    attachRowMeta(fakePkg(), [sheet])
    expect(sheet.rows.get(0)?.customHeight).toBeUndefined()
  })
})
