/**
 * 解析 worksheet 的 extLst → x14:sparklineGroups,回填到 SheetModel.sparklines。
 * ExcelJS 不解析迷你图,这里读原始 XML。
 *
 * 结构(removeNSPrefix 后 x14:/xm: 前缀已去掉):
 *   worksheet > extLst > ext[] > sparklineGroups > sparklineGroup[]
 *     @_type: 缺省=line / "column" / "stacked"(=winLoss)
 *     colorSeries/@_rgb, colorNegative/@_rgb
 *     sparklines > sparkline[]: f(数据区域 A1) + sqref(位置单元格)
 */
import type { RawPackage } from './raw-xml'
import { parseRels, toArray } from './raw-xml'
import type { CssColor, SheetModel, Sparkline } from '../model/types'
import { cellKey } from '../model/types'
import { parseA1Range } from './exceljs-adapter'

export function attachSparklines(pkg: RawPackage, sheets: SheetModel[]): void {
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
    const path = nameToPath.get(sheet.name)
    if (!path) continue
    try {
      parseSheetSparklines(pkg, path, sheet)
    } catch {
      /* 单个 sheet 解析失败不影响其它 */
    }
  }
}

function parseSheetSparklines(pkg: RawPackage, sheetPath: string, sheet: SheetModel): void {
  const xml = pkg.parse(sheetPath)
  const exts = toArray(xml?.worksheet?.extLst?.ext)
  for (const ext of exts) {
    const groups = toArray(ext.sparklineGroups?.sparklineGroup)
    for (const group of groups) {
      const typeAttr = String(group['@_type'] ?? 'line')
      const type: Sparkline['type'] =
        typeAttr === 'column' ? 'column' : typeAttr === 'stacked' ? 'winloss' : 'line'
      const color = rgbAttr(group.colorSeries)
      const negativeColor = rgbAttr(group.colorNegative)

      for (const sp of toArray(group.sparklines?.sparkline)) {
        const dataRef = textOf(sp.f)
        const loc = textOf(sp.sqref)
        if (!dataRef || !loc) continue
        const locCell = firstCellOf(loc)
        if (!locCell) continue
        const values = readRangeValues(sheet, dataRef)
        if (!values.length) continue
        sheet.sparklines.push({ row: locCell.row, col: locCell.col, type, values, color, negativeColor })
      }
    }
  }
}

/** 读 "Sheet1!C3:N3" 这种区域的数值(假定同 sheet);保留 null 作为缺口 */
function readRangeValues(sheet: SheetModel, ref: string): (number | null)[] {
  const a1 = ref.includes('!') ? ref.slice(ref.indexOf('!') + 1) : ref
  const rg = parseA1Range(a1.replace(/\$/g, ''))
  if (!rg) return []
  const out: (number | null)[] = []
  // 行向量或列向量都展开成一维序列
  for (let r = rg.top; r <= rg.bottom; r++) {
    for (let c = rg.left; c <= rg.right; c++) {
      const cell = sheet.cells.get(cellKey(r, c))
      out.push(cell && typeof cell.raw === 'number' ? cell.raw : null)
    }
  }
  return out
}

function firstCellOf(sqref: string): { row: number; col: number } | null {
  const first = sqref.trim().split(/\s+/)[0]
  const rg = parseA1Range(first)
  return rg ? { row: rg.top, col: rg.left } : null
}

function textOf(node: any): string | undefined {
  if (node == null) return undefined
  if (typeof node === 'string' || typeof node === 'number') return String(node)
  if (typeof node === 'object' && node['#text'] != null) return String(node['#text'])
  return undefined
}

function rgbAttr(node: any): CssColor | undefined {
  const rgb = node?.['@_rgb']
  if (!rgb) return undefined
  const hex = String(rgb)
  return '#' + (hex.length === 8 ? hex.slice(2) : hex).toUpperCase()
}
