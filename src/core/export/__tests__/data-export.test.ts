import { describe, it, expect } from 'vitest'
import { toCsv, toWorkbookJson } from '../data-export'
import { getSheetData } from '../../model/data-access'
import type { CellStyle, SheetModel, WorkbookModel } from '../../model/types'
import { cellKey } from '../../model/types'

const styleA = { font: {}, fill: { type: 'none' }, numFmt: 'General' } as unknown as CellStyle
function sheet(): SheetModel {
  const cells = new Map([
    [cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: '产品', styleId: 0 }],
    [cellKey(0, 1), { row: 0, col: 1, type: 'string', raw: '数量', styleId: 0 }],
    [cellKey(1, 0), { row: 1, col: 0, type: 'string', raw: 'a,b', styleId: 0 }], // 含逗号 → 需引号
    [cellKey(1, 1), { row: 1, col: 1, type: 'number', raw: 5, styleId: 0 }],
  ])
  return { name: 'S1', state: 'visible', cells, styles: [styleA], merges: [], dimension: { rows: 2, cols: 2 } } as unknown as SheetModel
}

describe('data-export CSV/JSON 与读层一致(E8:一份数据层)', () => {
  it('toCsv:与 getSheetData 同源,含逗号的格被引号包裹', () => {
    const s = sheet()
    const csv = toCsv(s)
    const lines = csv.split('\r\n')
    expect(lines[0]).toBe('产品,数量')
    expect(lines[1]).toBe('"a,b",5') // 逗号转义 + 数值
    // 与 getSheetData 行数一致
    expect(lines).toHaveLength(getSheetData(s).length)
  })

  it('toWorkbookJson:首行作 key 的对象数组', () => {
    const wb = { sheets: [sheet()], activeSheet: 0, themeColors: [], date1904: false } as unknown as WorkbookModel
    const json = JSON.parse(toWorkbookJson(wb))
    expect(json.S1).toEqual([{ 产品: 'a,b', 数量: 5 }])
  })
})
