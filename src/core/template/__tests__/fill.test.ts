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

describe('fillTemplate — 锚点 trimUnused(P3 进阶)', () => {
  // 构造模拟模板:5 行带边框的"空白数据区"(行 5-9)+ 合计行(行 10)
  function makeTemplate() {
    const wb = jsonToWorkbook([[]])
    const sheet = wb.sheets[0]
    sheet.dimension = { rows: 11, cols: 3 }
    // 模板的空白带边框格(模拟 6 行预留 placeholder 区)
    for (let r = 5; r < 10; r++) {
      for (let c = 0; c < 3; c++) {
        sheet.cells.set(cellKey(r, c), { row: r, col: c, type: 'empty', raw: null, styleId: 0 })
      }
    }
    // 模板的"真"内容:合计行(行 10),A 列写 '合计:',C 列写占位符
    sheet.cells.set(cellKey(10, 0), { row: 10, col: 0, type: 'string', raw: '合计:', styleId: 0 })
    sheet.cells.set(cellKey(10, 2), { row: 10, col: 2, type: 'string', raw: '{{total}}', styleId: 0 })
    return wb
  }

  // 注:'A6' = (row=5, col=0)(parseCellAddress 把 1-based 行号 -1)
  it('默认 trim:JSON 行少于模板预留 → 清空多余空白行;但不动后面有内容的合计行', async () => {
    const wb = makeTemplate()
    await fillTemplate(wb, {
      placeholders: { total: '1000' },
      anchors: [{ startCell: 'A6', rows: [['apple', 1, 100], ['banana', 2, 200]] }],
    })
    const sheet = wb.sheets[0]
    // 行 5-6 填了 JSON(0-based)= A6/A7
    expect(sheet.cells.get(cellKey(5, 0))?.raw).toBe('apple')
    expect(sheet.cells.get(cellKey(6, 0))?.raw).toBe('banana')
    // 行 7-9 应该被清空(JSON 只填 2 行,模板预留 rows 5-9 共 5 行)
    expect(sheet.cells.get(cellKey(7, 0))).toBeUndefined()
    expect(sheet.cells.get(cellKey(7, 2))).toBeUndefined()
    expect(sheet.cells.get(cellKey(9, 1))).toBeUndefined()
    // 行 10 (合计) 不动,{{total}} 已被占位符替换
    expect(sheet.cells.get(cellKey(10, 0))?.raw).toBe('合计:')
    expect(sheet.cells.get(cellKey(10, 2))?.raw).toBe('1000')
  })

  it('trimUnused: false → 保留模板原样,空白行内空格仍在', async () => {
    const wb = makeTemplate()
    await fillTemplate(wb, {
      anchors: [{ startCell: 'A6', rows: [['apple', 1, 100]], trimUnused: false }],
    })
    const sheet = wb.sheets[0]
    // 仅填 1 行;其余 4 行的空格保留(empty 占位)
    expect(sheet.cells.get(cellKey(5, 0))?.raw).toBe('apple')
    expect(sheet.cells.get(cellKey(7, 0))).toBeDefined() // 模板的边框/占位还在
    expect(sheet.cells.get(cellKey(9, 1))).toBeDefined()
  })

  it('JSON 行 >= 模板预留 → trim 没东西可清,行为不变', async () => {
    const wb = makeTemplate()
    await fillTemplate(wb, {
      placeholders: { total: '999' },
      anchors: [{ startCell: 'A6', rows: [
        ['a', 1, 10], ['b', 2, 20], ['c', 3, 30], ['d', 4, 40], ['e', 5, 50],
      ] }],
    })
    const sheet = wb.sheets[0]
    // 5 行都填了 → 最后一行(0-based row 9 = A10)
    expect(sheet.cells.get(cellKey(9, 0))?.raw).toBe('e')
    // 合计行不动
    expect(sheet.cells.get(cellKey(10, 0))?.raw).toBe('合计:')
  })

  it('只清"锚点列范围"内的格,不动锚点外的列', async () => {
    const wb = jsonToWorkbook([[]])
    const sheet = wb.sheets[0]
    sheet.dimension = { rows: 10, cols: 5 }
    // 行 5-9 的所有列(A-E)都放占位空格
    for (let r = 5; r < 10; r++) {
      for (let c = 0; c < 5; c++) {
        sheet.cells.set(cellKey(r, c), { row: r, col: c, type: 'empty', raw: null, styleId: 0 })
      }
    }
    await fillTemplate(wb, {
      // 锚点只用 A-C(3 列),D/E 是模板的其他内容
      anchors: [{ startCell: 'A6', rows: [['x', 1, 10]] }],
    })
    // 行 6-9 的 A-C 被清(锚点列范围)
    expect(sheet.cells.get(cellKey(6, 0))).toBeUndefined()
    expect(sheet.cells.get(cellKey(6, 2))).toBeUndefined()
    expect(sheet.cells.get(cellKey(9, 2))).toBeUndefined()
    // 行 6-9 的 D/E(锚点列外)不动
    expect(sheet.cells.get(cellKey(6, 3))).toBeDefined()
    expect(sheet.cells.get(cellKey(9, 4))).toBeDefined()
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
