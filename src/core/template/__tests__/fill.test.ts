import { describe, it, expect } from 'vitest'
import { fillTemplate, replacePlaceholders, parseCellAddress } from '../fill'
import { jsonToWorkbook } from '../../loader-json'
import { cellKey } from '../../model/types'

describe('parseCellAddress', () => {
  it('A1 → (0,0);AB12 → (11, 27);坏地址 → null', () => {
    expect(parseCellAddress('A1')).toEqual({ row: 0, col: 0 })
    expect(parseCellAddress('B3')).toEqual({ row: 2, col: 1 })
    expect(parseCellAddress('AB12')).toEqual({ row: 11, col: 27 })
    expect(parseCellAddress('xyz')).toBeNull()
  })
})

describe('replacePlaceholders', () => {
  it('{{key}} 替换;{{a.b.c}} dot path;缺失保留', () => {
    const ph = { name: 'Alice', user: { age: 30, addr: { city: 'NYC' } } }
    expect(replacePlaceholders('Hi {{name}}', ph)).toBe('Hi Alice')
    expect(replacePlaceholders('{{user.age}} 岁', ph)).toBe('30 岁')
    expect(replacePlaceholders('{{user.addr.city}}', ph)).toBe('NYC')
    expect(replacePlaceholders('{{missing}}', ph)).toBe('{{missing}}')
    expect(replacePlaceholders('混合 {{name}} 和 {{missing}}', ph)).toBe('混合 Alice 和 {{missing}}')
  })
})

describe('fillTemplate — 占位符', () => {
  it('扫全表 string 格,替换 {{key}};数字格不动', async () => {
    const wb = jsonToWorkbook([['客户: {{customer}}', 100]])
    await fillTemplate(wb, { placeholders: { customer: 'Alice' } })
    expect(wb.sheets[0].cells.get(cellKey(0, 0))?.raw).toBe('客户: Alice')
    expect(wb.sheets[0].cells.get(cellKey(0, 1))?.raw).toBe(100)
  })

  it('未匹配的占位符保留原样,不报错', async () => {
    const wb = jsonToWorkbook([['{{name}} - {{missing}}']])
    await fillTemplate(wb, { placeholders: { name: 'Bob' } })
    expect(wb.sheets[0].cells.get(cellKey(0, 0))?.raw).toBe('Bob - {{missing}}')
  })
})

describe('fillTemplate — 锚点表', () => {
  it('startCell 字符串 + rows 二维数组:从指定格起按位铺', async () => {
    const wb = jsonToWorkbook([['header']])
    await fillTemplate(wb, {
      anchors: [{ startCell: 'A3', rows: [['apple', 1], ['banana', 2]] }],
    })
    // A3 = row=2, col=0
    expect(wb.sheets[0].cells.get(cellKey(2, 0))?.raw).toBe('apple')
    expect(wb.sheets[0].cells.get(cellKey(2, 1))?.raw).toBe(1)
    expect(wb.sheets[0].cells.get(cellKey(3, 0))?.raw).toBe('banana')
    expect(wb.sheets[0].cells.get(cellKey(3, 1))?.raw).toBe(2)
  })

  it('对象数组:按 columns 顺序铺;不给 columns 用首行 keys', async () => {
    const wb = jsonToWorkbook([])
    await fillTemplate(wb, {
      anchors: [{
        startCell: { row: 0, col: 0 },
        rows: [{ name: 'a', age: 30 }, { name: 'b', age: 25 }],
        columns: ['age', 'name'], // 显式列序:age 在前
      }],
    })
    expect(wb.sheets[0].cells.get(cellKey(0, 0))?.raw).toBe(30)
    expect(wb.sheets[0].cells.get(cellKey(0, 1))?.raw).toBe('a')
    expect(wb.sheets[0].cells.get(cellKey(1, 0))?.raw).toBe(25)
  })

  it('空/null 值跳过(不覆盖模板原有值)', async () => {
    const wb = jsonToWorkbook([['template-value']])
    await fillTemplate(wb, {
      anchors: [{ startCell: 'A1', rows: [[null]] }],
    })
    // 锚点写 null 不动模板,A1 仍是 'template-value'
    expect(wb.sheets[0].cells.get(cellKey(0, 0))?.raw).toBe('template-value')
  })
})

describe('fillTemplate — 占位符 + 锚点组合 + 进度回调', () => {
  it('两者都跑;进度回调被调到至少一次', async () => {
    const wb = jsonToWorkbook([['Header: {{title}}']])
    let progressHits = 0
    const r = await fillTemplate(wb, {
      placeholders: { title: 'Invoice' },
      anchors: [{ startCell: 'A3', rows: [['x', 1]] }],
      onProgress: () => progressHits++,
    })
    expect(wb.sheets[0].cells.get(cellKey(0, 0))?.raw).toBe('Header: Invoice')
    expect(wb.sheets[0].cells.get(cellKey(2, 0))?.raw).toBe('x')
    expect(r.placeholdersScanned).toBeGreaterThanOrEqual(0)
    expect(r.anchorsWritten).toBe(2)
    expect(progressHits).toBeGreaterThanOrEqual(1)
  })

  it('pre-aborted signal:抛 AbortError', async () => {
    const wb = jsonToWorkbook([['Hi {{name}}']])
    const ctrl = new AbortController()
    ctrl.abort()
    try {
      await fillTemplate(wb, { placeholders: { name: 'X' }, signal: ctrl.signal })
      throw new Error('应当抛 AbortError')
    } catch (e) {
      expect((e as Error & { name: string }).name).toBe('AbortError')
    }
  })
})
