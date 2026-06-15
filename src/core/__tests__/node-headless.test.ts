/**
 * 纯 Node(headless)用法回归网 —— 锁住 1.15.0 给 Node 环境提供的能力:
 *  ① openWorkbook 直接吃 Node Buffer(F1 footgun);
 *  ② 取数 getSheetData / sheetToJSON / getCellText(显示文本走数字格式引擎);
 *  ③ 往返:setCellValue → workbookToXlsxBytes(Uint8Array,非 Blob)→ 重新解析改动保住;
 *  ④ 从数据建表:jsonToWorkbook → workbookToXlsxBytes → 解析回来。
 * 全程不碰 DOM / canvas。
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  openWorkbook,
  parseWorkbook,
  getSheetData,
  sheetToJSON,
  getCellText,
  setCellValue,
  workbookToXlsxBytes,
  jsonToWorkbook,
  cellKey,
} from '../index'

const SAMPLE = join(__dirname, '..', '..', '..', 'public', 'sample.xlsx')

describe('Node headless 用法', () => {
  it('openWorkbook 直接受理 Node Buffer(不必手动转 ArrayBuffer)', async () => {
    const buf = readFileSync(SAMPLE) // Node Buffer == Uint8Array 子类
    const wb = await openWorkbook(buf)
    expect(wb.sheets.length).toBe(3)
    expect(wb.sheets[0].name).toBe('销售报表')
  })

  it('parseWorkbook 也接受 Uint8Array(归一化)', async () => {
    const u8 = new Uint8Array(readFileSync(SAMPLE))
    const wb = await parseWorkbook(u8)
    expect(wb.sheets[0].cells.size).toBeGreaterThan(0)
  })

  it('取数: getSheetData / sheetToJSON / 显示文本', async () => {
    const wb = await openWorkbook(readFileSync(SAMPLE))
    const sheet = wb.sheets[0]

    const grid = getSheetData(sheet, { format: true })
    expect(Array.isArray(grid)).toBe(true)
    expect(grid.length).toBeGreaterThan(0)

    const rows = sheetToJSON(sheet)
    expect(Array.isArray(rows)).toBe(true)
    expect(rows.length).toBeGreaterThan(0)
    // 对象数组: 每行是 { 表头: 值 }
    expect(typeof rows[0]).toBe('object')

    // getCellText 给"人看到的字符串"
    const a1 = getCellText(sheet, 0, 0)
    expect(typeof a1).toBe('string')
    expect(a1.length).toBeGreaterThan(0)
  })

  it('高保真往返: setCellValue → workbookToXlsxBytes(Uint8Array)→ 重新解析改动保住', async () => {
    const src = readFileSync(SAMPLE)
    const wb = await openWorkbook(src)
    const sheet = wb.sheets[0]

    // 写一个原件里空的远端格(避开合并标题区),验证编辑能往返
    const R = 100
    const C = 0
    setCellValue(sheet, R, C, '__node_edited__')

    // overlay 高保真:保留原件 ExcelJS 能往返的部分
    const bytes = await workbookToXlsxBytes(wb, {
      fidelity: 'overlay',
      sourceBuffer: src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength),
    })
    // 关键: 返回的是 Uint8Array,不是浏览器 Blob
    expect(bytes).toBeInstanceOf(Uint8Array)
    expect(bytes.byteLength).toBeGreaterThan(0)

    const wb2 = await parseWorkbook(bytes)
    expect(wb2.sheets[0].cells.get(cellKey(R, C))?.raw).toBe('__node_edited__')
  })

  it('从数据建表: jsonToWorkbook → workbookToXlsxBytes → 解析回来', async () => {
    const wb = jsonToWorkbook(
      [
        { name: '张三', age: 25 },
        { name: '李四', age: 30 },
      ],
      { sheetName: 'People' },
    )
    const bytes = await workbookToXlsxBytes(wb)
    expect(bytes).toBeInstanceOf(Uint8Array)

    const wb2 = await parseWorkbook(bytes)
    const s = wb2.sheets[0]
    // 表头 + 2 行数据
    expect(getCellText(s, 0, 0)).toBe('name')
    expect(getCellText(s, 1, 0)).toBe('张三')
    expect(getCellText(s, 2, 0)).toBe('李四')
  })
})
