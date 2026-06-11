/**
 * 解析 worksheet 关联的 pivotTableDefinition + pivotCacheDefinition,
 * 回填只读透视表 UI 元数据。数据本身仍按普通单元格显示。
 */
import type { RawPackage } from './raw-xml'
import { parseRels, toArray } from './raw-xml'
import { parseA1Range } from './exceljs-adapter'
import type { MergeRange, PivotButton, PivotTableModel, SheetModel } from '../model/types'

export function attachPivotTables(pkg: RawPackage, sheets: SheetModel[]): void {
  const wbXml = pkg.parse('xl/workbook.xml')
  const sheetNodes = toArray(wbXml?.workbook?.sheets?.sheet)
  const wbRels = parseRels(pkg, 'xl/workbook.xml')
  const nameToPath = new Map<string, string>()
  for (const sn of sheetNodes) {
    const name = sn['@_name']
    const rid = sn['@_id']
    if (name && rid && wbRels[rid]) nameToPath.set(String(name), wbRels[rid])
  }

  for (const sheet of sheets) {
    const sheetPath = nameToPath.get(sheet.name)
    if (!sheetPath) continue
    try {
      attachSheetPivotTables(pkg, sheetPath, sheet)
    } catch {
      /* 单个 sheet 失败不影响其它 */
    }
  }
}

function attachSheetPivotTables(pkg: RawPackage, sheetPath: string, sheet: SheetModel): void {
  const sheetRels = parseRels(pkg, sheetPath)
  const paths = new Set<string>()
  // 标准 ECMA-376:worksheet 通过 rels 隐式关联 pivotTable 零件(sheet XML 里没有元素)
  for (const target of Object.values(sheetRels)) {
    if (/pivotTables\/pivotTable\d*\.xml$/i.test(target)) paths.add(target)
  }
  // 兼容:个别生成器在 worksheet XML 里留 pivotTableDefinition r:id 引用
  const sheetXml = pkg.parse(sheetPath)?.worksheet
  for (const ref of toArray(sheetXml?.pivotTableDefinition)) {
    const rid = ref['@_id']
    if (rid && sheetRels[rid]) paths.add(sheetRels[rid])
  }
  for (const pivotPath of paths) {
    const pivot = parsePivotTable(pkg, pivotPath, sheet)
    if (pivot) sheet.pivotTables.push(pivot)
  }
}

function parsePivotTable(pkg: RawPackage, pivotPath: string, sheet: SheetModel): PivotTableModel | null {
  const def = pkg.parse(pivotPath)?.pivotTableDefinition
  if (!def) return null
  const range = parseLocation(def.location)
  if (!range) return null
  const rels = parseRels(pkg, pivotPath)
  const cachePath = def['@_cacheId'] != null ? findCachePathById(pkg, String(def['@_cacheId'])) : undefined
  const relCachePath = Object.values(rels).find((p) => /pivotCache\/pivotCacheDefinition\d+\.xml$/i.test(p))
  const fields = readCacheFields(pkg, relCachePath ?? cachePath)
  const buttons = buildButtons(def, range, fields, sheet)
  return {
    name: String(def['@_name'] ?? def['@_pivotTableStyleInfo']?.['@_name'] ?? 'PivotTable'),
    range,
    fields,
    buttons,
  }
}

function parseLocation(loc: any): MergeRange | null {
  const ref = loc?.['@_ref']
  if (!ref) return null
  return parseA1Range(String(ref).replace(/\$/g, ''))
}

function readCacheFields(pkg: RawPackage, cachePath: string | undefined): string[] {
  if (!cachePath) return []
  const fields = toArray(pkg.parse(cachePath)?.pivotCacheDefinition?.cacheFields?.cacheField)
  return fields.map((f: any, i) => String(f?.['@_name'] ?? `字段${i + 1}`))
}

function findCachePathById(pkg: RawPackage, cacheId: string): string | undefined {
  const wb = pkg.parse('xl/workbook.xml')?.workbook
  const caches = toArray(wb?.pivotCaches?.pivotCache)
  const cache = caches.find((c: any) => String(c?.['@_cacheId']) === cacheId)
  const rid = cache?.['@_id']
  if (!rid) return undefined
  return parseRels(pkg, 'xl/workbook.xml')[rid]
}

function buildButtons(def: any, range: MergeRange, fields: string[], sheet: SheetModel): PivotButton[] {
  const buttons: PivotButton[] = []
  const used = new Set<string>()
  const add = (row: number, col: number, label: string, kind: PivotButton['kind']) => {
    if (row < 0 || col < 0) return
    if (row < range.top || row > range.bottom || col < range.left || col > range.right) return
    const key = `${row}:${col}:${label}:${kind}`
    if (used.has(key)) return
    used.add(key)
    buttons.push({ row, col, label, kind })
  }

  const pageFields = readFields(def.pageFields?.pageField, fields, '@_fld')
  pageFields.forEach((label, i) => add(range.top + i, range.left, label, 'page'))

  const colFields = readFields(def.colFields?.field, fields, '@_x')
  const rowFields = readFields(def.rowFields?.field, fields, '@_x')
  const dataFields = toArray(def.dataFields?.dataField).map((f: any, i) => String(f?.['@_name'] ?? fields[Number(f?.['@_fld'])] ?? `值${i + 1}`))

  const firstHeaderRow = clamp(range.top + Number(def.location?.['@_firstHeaderRow'] ?? 0), range.top, range.bottom)
  const firstDataCol = clamp(range.left + Number(def.location?.['@_firstDataCol'] ?? rowFields.length), range.left, range.right)
  const firstDataRow = clamp(range.top + Number(def.location?.['@_firstDataRow'] ?? colFields.length + 1), range.top, range.bottom)
  const rowHeaderRow = clamp(firstDataRow - 1, range.top, range.bottom)

  colFields.forEach((label, i) => add(firstHeaderRow + i, firstDataCol, label, 'col'))
  rowFields.forEach((label, i) => add(rowHeaderRow, range.left + i, label, 'row'))
  dataFields.forEach((label, i) => add(rowHeaderRow, Math.min(range.right, firstDataCol + i), label, 'data'))

  if (!buttons.length) {
    for (let col = range.left; col <= range.right; col++) {
      const text = cellText(sheet, range.top, col)
      if (text) add(range.top, col, text, 'field')
    }
  }
  return buttons
}

function readFields(nodes: any, fields: string[], attr: '@_x' | '@_fld'): string[] {
  return toArray(nodes)
    .map((f: any) => Number(f?.[attr]))
    .filter((n) => Number.isFinite(n) && n >= 0)
    .map((n) => fields[n] ?? `字段${n + 1}`)
}

function cellText(sheet: SheetModel, row: number, col: number): string {
  const raw = sheet.cells.get(`${row}:${col}`)?.raw
  return raw == null ? '' : String(raw)
}

function clamp(n: number, min: number, max: number): number {
  if (!Number.isFinite(n)) return min
  return Math.max(min, Math.min(max, n))
}
