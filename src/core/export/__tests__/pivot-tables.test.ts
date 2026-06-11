import { describe, it, expect } from 'vitest'
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate'
import { injectPivotTablesIntoZip, restoreOriginalPivotPartsIntoZip } from '../pivot-tables'
import { openPackage } from '../../parser/raw-xml'
import { attachPivotTables } from '../../parser/pivot-parser'
import { makeSheet } from '../../__tests__/helpers'
import type { CellModel, PivotTableLayout, SheetModel, WorkbookModel } from '../../model/types'
import { cellKey } from '../../model/types'

/** 模拟 ExcelJS 写出的最小 zip(workbook + 1 sheet + rels + content types)。 */
function fakeExcelJsZip(sheetNames = ['Sheet1']): Uint8Array {
  const files: Record<string, Uint8Array> = {}
  const put = (k: string, s: string) => {
    const d = strToU8(s)
    const c = new Uint8Array(d.byteLength)
    c.set(d)
    files[k] = c
  }
  const sheetTags = sheetNames.map((n, i) => `<sheet name="${n}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')
  const relTags = sheetNames.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')
  const overrides = sheetNames.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')
  put('[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    overrides + '</Types>')
  put('_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
  put('xl/workbook.xml',
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<sheets>${sheetTags}</sheets><calcPr calcId="0"/></workbook>`)
  put('xl/_rels/workbook.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relTags}</Relationships>`)
  sheetNames.forEach((_, i) => {
    put(`xl/worksheets/sheet${i + 1}.xml`,
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>')
  })
  return zipSync(files)
}

/** 源数据: A1:C5 表头 地区/产品/金额,4 行数据(含 1 个空白地区)。 */
function sourceSheet(): SheetModel {
  const sheet = makeSheet({ name: 'Sheet1' })
  const rows: Array<[string | number | null, string, number]> = [
    ['华东', '手机', 100],
    ['华南', '电脑', 200],
    ['华东', '电脑', 300],
    [null, '手机', 50],
  ]
  const putCell = (row: number, col: number, raw: string | number) => {
    const cell: CellModel = { row, col, type: typeof raw === 'number' ? 'number' : 'string', raw, styleId: 0 }
    sheet.cells.set(cellKey(row, col), cell)
  }
  putCell(0, 0, '地区')
  putCell(0, 1, '产品')
  putCell(0, 2, '金额')
  rows.forEach(([region, product, amount], i) => {
    if (region != null) putCell(i + 1, 0, region)
    putCell(i + 1, 1, product)
    putCell(i + 1, 2, amount)
  })
  return sheet
}

function makeWorkbook(layout: PivotTableLayout): WorkbookModel {
  const sheet = sourceSheet()
  sheet.pivotTables.push({
    name: 'PivotTable1',
    range: { top: 0, left: 5, bottom: 4, right: 7 },
    fields: ['地区', '产品', '金额'],
    buttons: [],
    source: { sheetIndex: 0, range: { top: 0, left: 0, bottom: 4, right: 2 } },
    layout,
  })
  return { sheets: [sheet], activeSheet: 0, themeColors: [], date1904: false }
}

const LAYOUT: PivotTableLayout = {
  filters: [{ field: 1, mode: 'all' }],
  columns: [],
  rows: [0],
  values: [{ field: 2, summary: 'sum' }],
}

describe('injectPivotTablesIntoZip', () => {
  it('回注全套标准 OOXML 零件(cache/records/table/rels/ContentTypes/workbook)', () => {
    const out = injectPivotTablesIntoZip(fakeExcelJsZip(), makeWorkbook(LAYOUT))
    const files = unzipSync(out)

    // 零件齐全
    expect(files['xl/pivotCache/pivotCacheDefinition1.xml']).toBeTruthy()
    expect(files['xl/pivotCache/pivotCacheRecords1.xml']).toBeTruthy()
    expect(files['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels']).toBeTruthy()
    expect(files['xl/pivotTables/pivotTable1.xml']).toBeTruthy()
    expect(files['xl/pivotTables/_rels/pivotTable1.xml.rels']).toBeTruthy()

    // cacheDefinition: 源区域 + 字段 + refreshOnLoad
    const cache = strFromU8(files['xl/pivotCache/pivotCacheDefinition1.xml'])
    expect(cache).toContain('<worksheetSource ref="A1:C5" sheet="Sheet1"/>')
    expect(cache).toContain('refreshOnLoad="1"')
    expect(cache).toContain('name="地区"')
    expect(cache).toContain('containsBlank="1"') // 地区列有空白
    expect(cache).toContain('containsNumber="1"') // 金额列数值 sharedItems
    expect(cache).toContain('recordCount="4"')

    // records: 4 行,轴字段写 <x>,值字段写 <n>,空白写 <m/>
    const records = strFromU8(files['xl/pivotCache/pivotCacheRecords1.xml'])
    expect(records).toContain('count="4"')
    expect(records).toContain('<n v="100"/>')
    expect(records).toContain('<m/>')

    // pivotTableDefinition: 行/页/值字段 + 样式
    const table = strFromU8(files['xl/pivotTables/pivotTable1.xml'])
    expect(table).toContain('<location ref="F1:H5"')
    expect(table).toContain('axis="axisRow"')
    expect(table).toContain('axis="axisPage"')
    expect(table).toContain('<rowFields count="1"><field x="0"/></rowFields>')
    expect(table).toContain('<pageFields count="1"><pageField fld="1" hier="-1"/></pageFields>')
    expect(table).toContain('fld="2"')
    expect(table).toContain('求和项:金额')
    expect(table).toContain('pivotTableStyleInfo')

    // workbook.xml + rels + ContentTypes 注册
    const wb = strFromU8(files['xl/workbook.xml'])
    expect(wb).toMatch(/<pivotCaches><pivotCache cacheId="\d+" r:id="rId\d+"\/><\/pivotCaches><\/workbook>/)
    const wbRels = strFromU8(files['xl/_rels/workbook.xml.rels'])
    expect(wbRels).toContain('pivotCache/pivotCacheDefinition1.xml')
    const ct = strFromU8(files['[Content_Types].xml'])
    expect(ct).toContain('/xl/pivotTables/pivotTable1.xml')
    expect(ct).toContain('pivotCacheDefinition+xml')
    // worksheet 隐式关系
    const sheetRels = strFromU8(files['xl/worksheets/_rels/sheet1.xml.rels'])
    expect(sheetRels).toContain('../pivotTables/pivotTable1.xml')
  })

  it('summary 非 sum 时写 subtotal 属性;列字段写 colFields', () => {
    const out = injectPivotTablesIntoZip(fakeExcelJsZip(), makeWorkbook({
      filters: [],
      columns: [1],
      rows: [0],
      values: [{ field: 2, summary: 'avg' }],
    }))
    const table = strFromU8(unzipSync(out)['xl/pivotTables/pivotTable1.xml'])
    expect(table).toContain('subtotal="average"')
    expect(table).toContain('平均值:金额')
    expect(table).toContain('<colFields count="1"><field x="1"/></colFields>')
    expect(table).toContain('axis="axisCol"')
  })

  it('无可回注透视表(缺 source/layout 元数据)时原样返回', () => {
    const zip = fakeExcelJsZip()
    const sheet = sourceSheet()
    sheet.pivotTables.push({ name: 'P', range: { top: 0, left: 5, bottom: 2, right: 6 }, fields: ['a'], buttons: [] })
    const wb: WorkbookModel = { sheets: [sheet], activeSheet: 0, themeColors: [], date1904: false }
    expect(injectPivotTablesIntoZip(zip, wb)).toBe(zip)
  })

  it('equals 筛选预选值 → pageField@item 指向选中项(打开还原筛选状态)', () => {
    const out = injectPivotTablesIntoZip(fakeExcelJsZip(), makeWorkbook({
      filters: [{ field: 1, mode: 'equals', value: '电脑' }],
      columns: [],
      rows: [0],
      values: [{ field: 2, summary: 'sum' }],
    }))
    const table = strFromU8(unzipSync(out)['xl/pivotTables/pivotTable1.xml'])
    // 产品列取值按出现序: 手机=0, 电脑=1 → 选中"电脑"写 item="1"
    expect(table).toContain('<pageField fld="1" item="1" hier="-1"/>')
  })

  it('include 多选筛选 → multipleItemSelectionAllowed + 未选项 item@h=1', () => {
    // 产品列(field 1)取值:手机(0)/电脑(1);include 只保留「手机」→ 电脑应隐藏
    const out = injectPivotTablesIntoZip(fakeExcelJsZip(), makeWorkbook({
      filters: [{ field: 1, mode: 'include', values: ['手机'] }],
      columns: [],
      rows: [0],
      values: [{ field: 2, summary: 'sum' }],
    }))
    const table = strFromU8(unzipSync(out)['xl/pivotTables/pivotTable1.xml'])
    expect(table).toContain('multipleItemSelectionAllowed="1"')
    expect(table).toContain('<item x="1" h="1"/>') // 电脑 = 索引 1,未选 → 隐藏
    expect(table).toContain('<item x="0"/>')        // 手机 = 索引 0,选中 → 不隐藏
    expect(table).toContain('<pageField fld="1" hier="-1"/>') // include 不写单选 item
  })

  it('non-empty 筛选 → 多选 + 隐藏空白项(WPS"去掉空白"语义)', () => {
    const out = injectPivotTablesIntoZip(fakeExcelJsZip(), makeWorkbook({
      filters: [{ field: 0, mode: 'non-empty' }], // 地区列含空白
      columns: [],
      rows: [1],
      values: [{ field: 2, summary: 'sum' }],
    }))
    const table = strFromU8(unzipSync(out)['xl/pivotTables/pivotTable1.xml'])
    expect(table).toContain('axis="axisPage" multipleItemSelectionAllowed="1"')
    expect(table).toContain('<item x="2" h="1"/>') // 华东=0 华南=1 空白=2 → 空白隐藏
    expect(table).toContain('<pageField fld="0" hier="-1"/>') // 多选不写 item
  })

  it('导出 → 重新解析:attachPivotTables 能在标准 rels 关联下发现回注的透视表(往返)', () => {
    const out = injectPivotTablesIntoZip(fakeExcelJsZip(), makeWorkbook(LAYOUT))
    const pkg = openPackage(out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength) as ArrayBuffer)
    const reparsed = makeSheet({ name: 'Sheet1' })
    attachPivotTables(pkg, [reparsed])
    expect(reparsed.pivotTables).toHaveLength(1)
    const pivot = reparsed.pivotTables[0]
    expect(pivot.name).toBe('PivotTable1')
    expect(pivot.range).toEqual({ top: 0, left: 5, bottom: 4, right: 7 })
    expect(pivot.fields).toEqual(['地区', '产品', '金额'])
    expect(pivot.buttons.length).toBeGreaterThan(0)
  })
})

describe('restoreOriginalPivotPartsIntoZip(overlay 保留原文件透视表)', () => {
  /** 模拟"原文件":带一套完整 pivot 零件(就用我们自己的回注产物当原件,结构同真实文件)。 */
  function originalWithPivot(): Uint8Array {
    return injectPivotTablesIntoZip(fakeExcelJsZip(), makeWorkbook(LAYOUT))
  }

  it('原件零件整套搬运 + workbook/sheet rels/ContentTypes 重新注册', () => {
    const restored = restoreOriginalPivotPartsIntoZip(fakeExcelJsZip(), originalWithPivot())
    const files = unzipSync(restored)
    expect(files['xl/pivotTables/pivotTable1.xml']).toBeTruthy()
    expect(files['xl/pivotCache/pivotCacheDefinition1.xml']).toBeTruthy()
    expect(files['xl/pivotCache/pivotCacheRecords1.xml']).toBeTruthy()
    expect(files['xl/pivotTables/_rels/pivotTable1.xml.rels']).toBeTruthy()
    expect(strFromU8(files['xl/workbook.xml'])).toMatch(/<pivotCaches><pivotCache cacheId="\d+" r:id="rId\d+"\/><\/pivotCaches>/)
    expect(strFromU8(files['xl/_rels/workbook.xml.rels'])).toContain('pivotCache/pivotCacheDefinition1.xml')
    expect(strFromU8(files['xl/worksheets/_rels/sheet1.xml.rels'])).toContain('../pivotTables/pivotTable1.xml')
    expect(strFromU8(files['[Content_Types].xml'])).toContain('/xl/pivotTables/pivotTable1.xml')
    // 搬运后可被我们自己的 parser 往返发现
    const pkg = openPackage(restored.buffer.slice(restored.byteOffset, restored.byteOffset + restored.byteLength) as ArrayBuffer)
    const sheet = makeSheet({ name: 'Sheet1' })
    attachPivotTables(pkg, [sheet])
    expect(sheet.pivotTables).toHaveLength(1)
  })

  it('搬运后再注 App 新建透视表:零件编号/cacheId 自动避开', () => {
    const restored = restoreOriginalPivotPartsIntoZip(fakeExcelJsZip(), originalWithPivot())
    const out = injectPivotTablesIntoZip(restored, makeWorkbook(LAYOUT))
    const files = unzipSync(out)
    expect(files['xl/pivotTables/pivotTable1.xml']).toBeTruthy() // 原件的
    expect(files['xl/pivotTables/pivotTable2.xml']).toBeTruthy() // App 新建的
    const wb = strFromU8(files['xl/workbook.xml'])
    const ids = [...wb.matchAll(/cacheId="(\d+)"/g)].map((m) => m[1])
    expect(new Set(ids).size).toBe(ids.length) // cacheId 不冲突
  })

  it('原件无透视表 / 原件损坏 → 原样返回', () => {
    const target = fakeExcelJsZip()
    expect(restoreOriginalPivotPartsIntoZip(target, fakeExcelJsZip())).toBe(target)
    expect(restoreOriginalPivotPartsIntoZip(target, strToU8('not a zip'))).toBe(target)
  })
})
