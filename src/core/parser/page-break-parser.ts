/**
 * 解析 worksheet 的手动分页符(rowBreaks/colBreaks),回填 SheetModel.pageBreaks。
 * ExcelJS 不暴露,读原始 XML。brk @_id = 分页边界的行/列(0-based 边界,在其上方/左侧断页)。
 */
import type { RawPackage } from './raw-xml'
import { parseRels, toArray } from './raw-xml'
import type { SheetModel } from '../model/types'

export function attachPageBreaks(pkg: RawPackage, sheets: SheetModel[]): void {
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
      const ws = pkg.parse(path)?.worksheet
      const rows = toArray(ws?.rowBreaks?.brk)
        .map((b: any) => Number(b['@_id']))
        .filter((n: number) => Number.isFinite(n) && n > 0)
      const cols = toArray(ws?.colBreaks?.brk)
        .map((b: any) => Number(b['@_id']))
        .filter((n: number) => Number.isFinite(n) && n > 0)
      if (rows.length || cols.length) sheet.pageBreaks = { rows, cols }
    } catch {
      /* 单 sheet 失败不影响其它 */
    }
  }
}
