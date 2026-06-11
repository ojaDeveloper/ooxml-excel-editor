/**
 * 真实 OOXML 透视表导出回注 —— 解决"App 内创建的透视表导出后只是普通单元格"。
 *
 * ExcelJS 不建模 pivot 零件(pivotCacheDefinition / pivotCacheRecords / pivotTableDefinition),
 * load→write 会整套丢掉。本模块在 ExcelJS 写出 zip **之后**做后处理(同 wps-cellimages.ts 模式),
 * 把 App 内创建的透视表(`PivotTableModel` 带 `source`+`layout` 元数据)重建成标准 ECMA-376 零件:
 *
 *   - xl/pivotCache/pivotCacheDefinition{n}.xml   缓存定义(cacheSource + cacheFields/sharedItems)
 *   - xl/pivotCache/pivotCacheRecords{n}.xml      缓存记录(源数据行)
 *   - xl/pivotCache/_rels/...                     definition → records
 *   - xl/pivotTables/pivotTable{n}.xml            透视表定义(location/pivotFields/row/col/page/dataFields)
 *   - xl/pivotTables/_rels/...                    pivotTable → cacheDefinition
 *   - xl/_rels/workbook.xml.rels                  + pivotCacheDefinition 关系
 *   - xl/workbook.xml                             + <pivotCaches><pivotCache cacheId r:id/></pivotCaches>
 *   - xl/worksheets/_rels/sheetN.xml.rels         + pivotTable 隐式关系(标准 OOXML 靠 rels 关联)
 *   - [Content_Types].xml                         + 三类零件 Override
 *
 * `refreshOnLoad="1"`:Excel/WPS 打开时自动按缓存源区域重算原生透视布局 —— 静态汇总结果仍写在
 * 单元格里(不打开透视功能的查看器也能看),刷新后被原生透视渲染替换。
 *
 * 筛选器导出语义(对齐 WPS/Excel 行为):`equals` 写 `pageField@item` 指向选中项 → 打开还原选中值;
 * `non-empty` 映射为多选 + 隐藏空白项(`multipleItemSelectionAllowed` + item@h)→ 打开显示"(多项)";
 * `all` 不写选中 → "全部"。
 *
 * 原文件已有的透视表(解析为只读,无 source/layout 元数据)无法从模型重建 —— overlay 导出时由
 * `restoreOriginalPivotPartsIntoZip` 从原件 zip **原样搬运**整套零件(见下),rebuild 模式因
 * 结构可能被增删行列改动,不搬运(退化为普通单元格)。
 */
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate'
import type { PivotSummary, PivotTableLayout, PivotTableModel, SheetModel, WorkbookModel } from '../model/types'
import { cellKey, type MergeRange } from '../model/types'

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const REL_PIVOT_CACHE = `${NS_REL}/pivotCacheDefinition`
const REL_PIVOT_RECORDS = `${NS_REL}/pivotCacheRecords`
const REL_PIVOT_TABLE = `${NS_REL}/pivotTable`
const CT_PIVOT_CACHE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'
const CT_PIVOT_RECORDS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml'
const CT_PIVOT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml'

interface PivotJob {
  pivot: PivotTableModel
  source: { sheetIndex: number; range: MergeRange }
  layout: PivotTableLayout
  hostSheet: SheetModel
  sourceSheet: SheetModel
}

/** 单字段的缓存形态:轴字段枚举 sharedItems(字符串);纯值字段写数值统计。 */
interface FieldCache {
  name: string
  onAxis: boolean
  numeric: boolean
  hasBlank: boolean
  /** 轴字段的去重取值(不含空白;空白用 hasBlank + <m/> 表达) */
  items: string[]
  /** 轴字段:取值 → sharedItems index(空白排最后) */
  itemIndex: Map<string, number>
  min: number
  max: number
}

function xmlEscape(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function colLabel(col: number): string {
  let n = col + 1
  let s = ''
  while (n > 0) {
    const r = (n - 1) % 26
    s = String.fromCharCode(65 + r) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

function rangeRef(r: MergeRange): string {
  return `${colLabel(r.left)}${r.top + 1}:${colLabel(r.right)}${r.bottom + 1}`
}

function put(files: Record<string, Uint8Array>, key: string, text: string): void {
  const data = strToU8(text)
  const copy = new Uint8Array(data.byteLength)
  copy.set(data)
  files[key] = copy
}

function rawAt(sheet: SheetModel, row: number, col: number): string | number | boolean | Date | null {
  const raw = sheet.cells.get(cellKey(row, col))?.raw
  return raw == null ? null : (raw as string | number | boolean | Date)
}

function rawText(raw: string | number | boolean | Date | null): string {
  if (raw == null) return ''
  if (raw instanceof Date) return raw.toISOString()
  return String(raw)
}

function rawNumber(raw: string | number | boolean | Date | null): number | null {
  if (typeof raw === 'number' && Number.isFinite(raw)) return raw
  if (typeof raw === 'string') {
    const n = Number(raw.replace(/,/g, '').trim())
    if (raw.trim() !== '' && Number.isFinite(n)) return n
  }
  return null
}

const SUMMARY_SUBTOTAL: Record<PivotSummary, string | undefined> = {
  sum: undefined, // sum 是 dataField@subtotal 缺省值,省略保持文件精简
  count: 'count',
  avg: 'average',
  max: 'max',
  min: 'min',
}

const SUMMARY_LABEL: Record<PivotSummary, string> = { sum: '求和项', count: '计数项', avg: '平均值', max: '最大值', min: '最小值' }

/** 收集需要回注的透视表(App 内创建、带完整元数据、源表仍存在)。 */
function collectJobs(workbook: WorkbookModel): PivotJob[] {
  const jobs: PivotJob[] = []
  for (const sheet of workbook.sheets) {
    for (const pivot of sheet.pivotTables ?? []) {
      const source = pivot.source
      const layout = pivot.layout
      if (!source || !layout || !layout.values.length) continue
      const sourceSheet = workbook.sheets[source.sheetIndex]
      if (!sourceSheet) continue
      if (source.range.bottom <= source.range.top) continue
      jobs.push({ pivot, source, layout, hostSheet: sheet, sourceSheet })
    }
  }
  return jobs
}

/** 逐字段算缓存形态(轴字段枚举取值,值字段统计 min/max)。 */
function buildFieldCaches(job: PivotJob): FieldCache[] {
  const { source, layout, sourceSheet, pivot } = job
  const axisCols = new Set<number>([...layout.rows, ...layout.columns, ...layout.filters.map((f) => f.field)])
  const out: FieldCache[] = []
  for (let col = source.range.left; col <= source.range.right; col++) {
    const i = col - source.range.left
    const fc: FieldCache = {
      name: pivot.fields[i] || `字段${i + 1}`,
      onAxis: axisCols.has(col),
      numeric: true,
      hasBlank: false,
      items: [],
      itemIndex: new Map(),
      min: Infinity,
      max: -Infinity,
    }
    const seen = new Set<string>()
    let sawValue = false
    for (let row = source.range.top + 1; row <= source.range.bottom; row++) {
      const raw = rawAt(sourceSheet, row, col)
      const text = rawText(raw).trim()
      if (!text) {
        fc.hasBlank = true
        continue
      }
      sawValue = true
      const n = rawNumber(raw)
      if (n == null) fc.numeric = false
      else {
        fc.min = Math.min(fc.min, n)
        fc.max = Math.max(fc.max, n)
      }
      if (fc.onAxis && !seen.has(text)) {
        seen.add(text)
        fc.itemIndex.set(text, fc.items.length)
        fc.items.push(text)
      }
    }
    if (!sawValue) fc.numeric = false
    out.push(fc)
  }
  return out
}

function cacheDefinitionXml(job: PivotJob, fields: FieldCache[], recordCount: number): string {
  const sourceName = job.sourceSheet.name
  const fieldXml = fields.map((fc) => {
    if (fc.onAxis) {
      const items = fc.items.map((v) => `<s v="${xmlEscape(v)}"/>`).join('') + (fc.hasBlank ? '<m/>' : '')
      const count = fc.items.length + (fc.hasBlank ? 1 : 0)
      const blankAttr = fc.hasBlank ? ' containsBlank="1"' : ''
      return `<cacheField name="${xmlEscape(fc.name)}" numFmtId="0"><sharedItems${blankAttr} count="${count}">${items}</sharedItems></cacheField>`
    }
    if (fc.numeric) {
      const blankAttr = fc.hasBlank ? ' containsBlank="1"' : ''
      return (
        `<cacheField name="${xmlEscape(fc.name)}" numFmtId="0">` +
        `<sharedItems${blankAttr} containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="${fc.min}" maxValue="${fc.max}"/></cacheField>`
      )
    }
    return `<cacheField name="${xmlEscape(fc.name)}" numFmtId="0"><sharedItems/></cacheField>`
  }).join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_REL}" r:id="rId1" refreshedBy="ooxml-excel-editor"` +
    ' refreshOnLoad="1" createdVersion="3" refreshedVersion="3" minRefreshableVersion="3"' +
    ` recordCount="${recordCount}">` +
    `<cacheSource type="worksheet"><worksheetSource ref="${rangeRef(job.source.range)}" sheet="${xmlEscape(sourceName)}"/></cacheSource>` +
    `<cacheFields count="${fields.length}">${fieldXml}</cacheFields>` +
    '</pivotCacheDefinition>'
  )
}

function cacheRecordsXml(job: PivotJob, fields: FieldCache[]): { xml: string; count: number } {
  const { source, sourceSheet } = job
  const lines: string[] = []
  for (let row = source.range.top + 1; row <= source.range.bottom; row++) {
    const cells = fields.map((fc, i) => {
      const raw = rawAt(sourceSheet, row, source.range.left + i)
      const text = rawText(raw).trim()
      if (!text) return '<m/>'
      if (fc.onAxis) return `<x v="${fc.itemIndex.get(text) ?? 0}"/>`
      const n = rawNumber(raw)
      if (fc.numeric && n != null) return `<n v="${n}"/>`
      return `<s v="${xmlEscape(text)}"/>`
    })
    lines.push(`<r>${cells.join('')}</r>`)
  }
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<pivotCacheRecords xmlns="${NS_MAIN}" xmlns:r="${NS_REL}" count="${lines.length}">` +
    lines.join('') +
    '</pivotCacheRecords>'
  return { xml, count: lines.length }
}

function pivotTableXml(job: PivotJob, fields: FieldCache[], cacheId: number): string {
  const { pivot, source, layout } = job
  const rel = (col: number) => col - source.range.left
  const axisOf = (i: number): string | null => {
    const col = source.range.left + i
    if (layout.rows.includes(col)) return 'axisRow'
    if (layout.columns.includes(col)) return 'axisCol'
    if (layout.filters.some((f) => f.field === col)) return 'axisPage'
    return null
  }
  const valueFields = new Set(layout.values.map((v) => rel(v.field)))
  const filterRuleOf = (i: number) => layout.filters.find((f) => rel(f.field) === i)
  const pivotFields = fields.map((fc, i) => {
    const axis = axisOf(i)
    if (axis) {
      const count = fc.items.length + (fc.hasBlank ? 1 : 0)
      const rule = axis === 'axisPage' ? filterRuleOf(i) : undefined
      const blankIdx = fc.items.length
      // 多选(include)/非空(non-empty)→ multipleItemSelectionAllowed + 未选项 item@h=1(WPS"勾选筛选"语义);
      // equals 的单选由 pageField@item 表达,不在此隐藏。
      const hidden = new Set<number>()
      let multi = false
      if (rule?.mode === 'non-empty' && fc.hasBlank) { multi = true; hidden.add(blankIdx) }
      if (rule?.mode === 'include' && rule.values) {
        multi = true
        const keep = new Set(rule.values)
        fc.items.forEach((v, idx) => { if (!keep.has(v)) hidden.add(idx) })
        if (fc.hasBlank) hidden.add(blankIdx) // 缓存项不含空白,include 永不选中空白 → 隐藏
      }
      const items = Array.from({ length: count }, (_, j) => hidden.has(j) ? `<item x="${j}" h="1"/>` : `<item x="${j}"/>`).join('') + '<item t="default"/>'
      const multiAttr = multi ? ' multipleItemSelectionAllowed="1"' : ''
      return `<pivotField axis="${axis}"${multiAttr} showAll="0"><items count="${count + 1}">${items}</items></pivotField>`
    }
    if (valueFields.has(i)) return '<pivotField dataField="1" showAll="0"/>'
    return '<pivotField showAll="0"/>'
  }).join('')

  const rowFields = layout.rows.length
    ? `<rowFields count="${layout.rows.length}">${layout.rows.map((col) => `<field x="${rel(col)}"/>`).join('')}</rowFields>`
    : ''
  const colFields = layout.columns.length
    ? `<colFields count="${layout.columns.length}">${layout.columns.map((col) => `<field x="${rel(col)}"/>`).join('')}</colFields>`
    : ''
  const pageFields = layout.filters.length
    ? `<pageFields count="${layout.filters.length}">${layout.filters.map((f) => {
        // equals 且值在缓存项里 → item 指向选中项索引,打开文件还原筛选状态(同 WPS 保存行为)
        const fc = fields[rel(f.field)]
        const idx = f.mode === 'equals' && f.value != null ? fc?.itemIndex.get(f.value) : undefined
        return `<pageField fld="${rel(f.field)}"${idx != null ? ` item="${idx}"` : ''} hier="-1"/>`
      }).join('')}</pageFields>`
    : ''
  const dataFields = `<dataFields count="${layout.values.length}">${layout.values.map((v) => {
    const fc = fields[rel(v.field)]
    const subtotal = SUMMARY_SUBTOTAL[v.summary]
    const name = `${SUMMARY_LABEL[v.summary]}:${fc?.name ?? `字段${rel(v.field) + 1}`}`
    return `<dataField name="${xmlEscape(name)}" fld="${rel(v.field)}" baseField="0" baseItem="0"${subtotal ? ` subtotal="${subtotal}"` : ''}/>`
  }).join('')}</dataFields>`

  const firstDataRow = layout.columns.length ? 2 : 1
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<pivotTableDefinition xmlns="${NS_MAIN}" name="${xmlEscape(pivot.name)}" cacheId="${cacheId}"` +
    ' applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0"' +
    ' applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="值"' +
    ' updatedVersion="3" minRefreshableVersion="3" createdVersion="3" useAutoFormatting="1"' +
    ' itemPrintTitles="1" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">' +
    `<location ref="${rangeRef(pivot.range)}" firstHeaderRow="1" firstDataRow="${firstDataRow}" firstDataCol="1"/>` +
    `<pivotFields count="${fields.length}">${pivotFields}</pivotFields>` +
    rowFields +
    colFields +
    pageFields +
    dataFields +
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>' +
    '</pivotTableDefinition>'
  )
}

const RELS_EMPTY =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

function appendRel(rels: string, type: string, target: string): { rels: string; rid: string } {
  let max = 0
  for (const m of rels.matchAll(/Id="rId(\d+)"/g)) max = Math.max(max, Number(m[1]))
  const rid = `rId${max + 1}`
  return {
    rels: rels.replace('</Relationships>', `<Relationship Id="${rid}" Type="${type}" Target="${target}"/></Relationships>`),
    rid,
  }
}

/** workbook.xml 里 sheet name → worksheet 零件路径(经 workbook.xml.rels)。 */
function sheetPathByName(files: Record<string, Uint8Array>, name: string): string | null {
  const wbXml = files['xl/workbook.xml'] ? strFromU8(files['xl/workbook.xml']) : ''
  const wbRels = files['xl/_rels/workbook.xml.rels'] ? strFromU8(files['xl/_rels/workbook.xml.rels']) : ''
  const escaped = xmlEscape(name)
  for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) {
    const tag = m[0]
    const nm = /name="([^"]*)"/.exec(tag)?.[1]
    if (nm !== escaped && nm !== name) continue
    const rid = /r:id="([^"]+)"/.exec(tag)?.[1]
    if (!rid) return null
    const rel = new RegExp(`<Relationship[^>]*Id="${rid}"[^>]*Target="([^"]+)"`).exec(wbRels)
    if (!rel) return null
    const target = rel[1]
    return target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`
  }
  return null
}

/**
 * 把 App 内创建的透视表回注进 ExcelJS 写出的 zip 字节,返回新 zip 字节。
 * 无可回注透视表 → 原样返回(零开销)。单个透视表失败不影响其余(整体 try/catch 由调用方兜底)。
 */
export function injectPivotTablesIntoZip(zipBytes: Uint8Array, workbook: WorkbookModel): Uint8Array {
  const jobs = collectJobs(workbook)
  if (!jobs.length) return zipBytes

  const files: Record<string, Uint8Array> = unzipSync(zipBytes)
  if (!files['xl/workbook.xml'] || !files['xl/_rels/workbook.xml.rels']) return zipBytes

  // 零件起始编号/缓存 id 避开 zip 里已有的(overlay 模式原件理论上已被 ExcelJS 丢弃,稳妥起见仍扫描)
  let partNo = 1
  for (const key of Object.keys(files)) {
    const m = /^xl\/pivotTables\/pivotTable(\d+)\.xml$/.exec(key)
    if (m) partNo = Math.max(partNo, Number(m[1]) + 1)
  }
  let wbXml = strFromU8(files['xl/workbook.xml'])
  let cacheId = 1
  for (const m of wbXml.matchAll(/cacheId="(\d+)"/g)) cacheId = Math.max(cacheId, Number(m[1]) + 1)

  let wbRels = strFromU8(files['xl/_rels/workbook.xml.rels'])
  let ctXml = files['[Content_Types].xml'] ? strFromU8(files['[Content_Types].xml']) : ''
  const pivotCacheTags: string[] = []

  for (const job of jobs) {
    const sheetPath = sheetPathByName(files, job.hostSheet.name)
    if (!sheetPath || !files[sheetPath]) continue
    const n = partNo++
    const id = cacheId++
    const fields = buildFieldCaches(job)
    const records = cacheRecordsXml(job, fields)

    put(files, `xl/pivotCache/pivotCacheDefinition${n}.xml`, cacheDefinitionXml(job, fields, records.count))
    put(files, `xl/pivotCache/pivotCacheRecords${n}.xml`, records.xml)
    put(
      files,
      `xl/pivotCache/_rels/pivotCacheDefinition${n}.xml.rels`,
      RELS_EMPTY.replace('</Relationships>', `<Relationship Id="rId1" Type="${REL_PIVOT_RECORDS}" Target="pivotCacheRecords${n}.xml"/></Relationships>`),
    )
    put(files, `xl/pivotTables/pivotTable${n}.xml`, pivotTableXml(job, fields, id))
    put(
      files,
      `xl/pivotTables/_rels/pivotTable${n}.xml.rels`,
      RELS_EMPTY.replace('</Relationships>', `<Relationship Id="rId1" Type="${REL_PIVOT_CACHE}" Target="../pivotCache/pivotCacheDefinition${n}.xml"/></Relationships>`),
    )

    // workbook rels + pivotCaches 注册
    const wb = appendRel(wbRels, REL_PIVOT_CACHE, `pivotCache/pivotCacheDefinition${n}.xml`)
    wbRels = wb.rels
    pivotCacheTags.push(`<pivotCache cacheId="${id}" r:id="${wb.rid}"/>`)

    // worksheet 隐式关系(标准 OOXML:透视表零件由所在 sheet 的 rels 关联,sheet XML 无元素)
    const dir = sheetPath.slice(0, sheetPath.lastIndexOf('/'))
    const base = sheetPath.slice(sheetPath.lastIndexOf('/') + 1)
    const sheetRelsKey = `${dir}/_rels/${base}.rels`
    const sheetRels = files[sheetRelsKey] ? strFromU8(files[sheetRelsKey]) : RELS_EMPTY
    put(files, sheetRelsKey, appendRel(sheetRels, REL_PIVOT_TABLE, `../pivotTables/pivotTable${n}.xml`).rels)

    // [Content_Types].xml Override
    let inserts = ''
    inserts += `<Override PartName="/xl/pivotCache/pivotCacheDefinition${n}.xml" ContentType="${CT_PIVOT_CACHE}"/>`
    inserts += `<Override PartName="/xl/pivotCache/pivotCacheRecords${n}.xml" ContentType="${CT_PIVOT_RECORDS}"/>`
    inserts += `<Override PartName="/xl/pivotTables/pivotTable${n}.xml" ContentType="${CT_PIVOT_TABLE}"/>`
    ctXml = ctXml.replace('</Types>', inserts + '</Types>')
  }

  if (!pivotCacheTags.length) return zipBytes

  put(files, 'xl/workbook.xml', insertPivotCaches(wbXml, pivotCacheTags))
  put(files, 'xl/_rels/workbook.xml.rels', wbRels)
  if (ctXml) put(files, '[Content_Types].xml', ctXml)

  return zipSync(files)
}

/** workbook.xml 插入 <pivotCache> 注册:并入已有 <pivotCaches>;否则按 CT_Workbook 序列放 extLst 前 / </workbook> 前。 */
function insertPivotCaches(wbXml: string, tags: string[]): string {
  if (wbXml.includes('<pivotCaches>')) return wbXml.replace('<pivotCaches>', `<pivotCaches>${tags.join('')}`)
  const caches = `<pivotCaches>${tags.join('')}</pivotCaches>`
  if (wbXml.includes('<extLst')) return wbXml.replace('<extLst', `${caches}<extLst`)
  return wbXml.replace('</workbook>', `${caches}</workbook>`)
}

const CT_BY_PART: Array<[RegExp, string]> = [
  [/pivotCacheDefinition\d*\.xml$/, CT_PIVOT_CACHE],
  [/pivotCacheRecords\d*\.xml$/, CT_PIVOT_RECORDS],
  [/pivotTables\/pivotTable\d*\.xml$/, CT_PIVOT_TABLE],
]

/**
 * overlay 导出:把**原文件**的透视表零件原样搬运进 ExcelJS 写出的 zip(ExcelJS 不建模 pivot,
 * load→write 会整套丢掉)。与 `injectPivotTablesIntoZip`(重建 App 内新建的)互补,先搬运后重建,
 * 重建侧的零件编号/cacheId 扫描会自动避开搬运进来的。
 *
 * 搬运内容:xl/pivotCache/** + xl/pivotTables/**(零件 + _rels 整目录)、workbook `<pivotCaches>`
 * 注册(cacheId 不变,r:id 在新 rels 里重新分配)、所在 worksheet 的隐式关系(按表名匹配新旧
 * worksheet)、`[Content_Types].xml` Override。源数据被编辑过时,打开后由宿主刷新重算(透视表
 * 自身带 cacheSource 范围)。原件无透视表 → 原样返回(零开销)。
 */
export function restoreOriginalPivotPartsIntoZip(zipBytes: Uint8Array, sourceBytes: Uint8Array): Uint8Array {
  let src: Record<string, Uint8Array>
  try {
    src = unzipSync(sourceBytes)
  } catch {
    return zipBytes
  }
  if (!src['xl/workbook.xml']) return zipBytes
  const srcWbXml = strFromU8(src['xl/workbook.xml'])
  const srcWbRels = src['xl/_rels/workbook.xml.rels'] ? strFromU8(src['xl/_rels/workbook.xml.rels']) : ''
  const caches = [...srcWbXml.matchAll(/<pivotCache\b[^>]*cacheId="(\d+)"[^>]*r:id="([^"]+)"[^>]*\/>/g)]
  const partKeys = Object.keys(src).filter((k) => k.startsWith('xl/pivotCache/') || k.startsWith('xl/pivotTables/'))
  if (!caches.length || !partKeys.length) return zipBytes

  const files: Record<string, Uint8Array> = unzipSync(zipBytes)
  if (!files['xl/workbook.xml'] || !files['xl/_rels/workbook.xml.rels']) return zipBytes

  // 1. 零件整目录搬运(已存在同名零件则跳过,不覆盖)
  for (const key of partKeys) {
    if (files[key]) continue
    const copy = new Uint8Array(src[key].byteLength)
    copy.set(src[key])
    files[key] = copy
  }

  // 2. workbook 注册:原 cacheId 保留,r:id 在新 workbook rels 里重新分配
  let wbRels = strFromU8(files['xl/_rels/workbook.xml.rels'])
  const tags: string[] = []
  for (const [, cacheId, srcRid] of caches) {
    const target = new RegExp(`<Relationship[^>]*Id="${srcRid}"[^>]*Target="([^"]+)"`).exec(srcWbRels)?.[1]
    if (!target) continue
    const normalized = target.startsWith('/') ? target.slice(1).replace(/^xl\//, '') : target.replace(/^\.\//, '')
    if (!files[`xl/${normalized}`]) continue
    const added = appendRel(wbRels, REL_PIVOT_CACHE, normalized)
    wbRels = added.rels
    tags.push(`<pivotCache cacheId="${cacheId}" r:id="${added.rid}"/>`)
  }
  if (!tags.length) return zipBytes
  put(files, 'xl/workbook.xml', insertPivotCaches(strFromU8(files['xl/workbook.xml']), tags))
  put(files, 'xl/_rels/workbook.xml.rels', wbRels)

  // 3. worksheet 隐式关系:原件里哪张表挂了哪些 pivotTable,按表名搬到新 zip 的对应 sheet rels
  const srcSheetName = new Map<string, string>() // 原 worksheet 路径 → 表名(XML 转义形态)
  for (const m of srcWbXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = /name="([^"]*)"/.exec(m[0])?.[1]
    const rid = /r:id="([^"]+)"/.exec(m[0])?.[1]
    if (!name || !rid) continue
    const target = new RegExp(`<Relationship[^>]*Id="${rid}"[^>]*Target="([^"]+)"`).exec(srcWbRels)?.[1]
    if (target) srcSheetName.set(target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`, name)
  }
  for (const [srcSheetPath, name] of srcSheetName) {
    const dir = srcSheetPath.slice(0, srcSheetPath.lastIndexOf('/'))
    const base = srcSheetPath.slice(srcSheetPath.lastIndexOf('/') + 1)
    const srcRelsBytes = src[`${dir}/_rels/${base}.rels`]
    if (!srcRelsBytes) continue
    const pivotTargets = [...strFromU8(srcRelsBytes).matchAll(/<Relationship[^>]*Type="[^"]*\/pivotTable"[^>]*Target="([^"]+)"/g)].map((m) => m[1])
    if (!pivotTargets.length) continue
    const newSheetPath = sheetPathByName(files, name)
    if (!newSheetPath) continue
    const newDir = newSheetPath.slice(0, newSheetPath.lastIndexOf('/'))
    const newBase = newSheetPath.slice(newSheetPath.lastIndexOf('/') + 1)
    const relsKey = `${newDir}/_rels/${newBase}.rels`
    let rels = files[relsKey] ? strFromU8(files[relsKey]) : RELS_EMPTY
    for (const target of pivotTargets) {
      if (rels.includes(`Target="${target}"`)) continue
      rels = appendRel(rels, REL_PIVOT_TABLE, target).rels
    }
    put(files, relsKey, rels)
  }

  // 4. Content_Types Override(按零件路径模式补)
  if (files['[Content_Types].xml']) {
    let ct = strFromU8(files['[Content_Types].xml'])
    let inserts = ''
    for (const key of partKeys) {
      if (key.includes('/_rels/')) continue
      const ctType = CT_BY_PART.find(([re]) => re.test(key))?.[1]
      if (ctType && !ct.includes(`/${key}`)) inserts += `<Override PartName="/${key}" ContentType="${ctType}"/>`
    }
    if (inserts) ct = ct.replace('</Types>', inserts + '</Types>')
    put(files, '[Content_Types].xml', ct)
  }

  return zipSync(files)
}
