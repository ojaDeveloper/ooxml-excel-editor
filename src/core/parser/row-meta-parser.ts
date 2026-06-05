/**
 * 从原始 sheetN.xml 读 `<row customHeight="1">` 标记,回填到 SheetModel.rows[r].customHeight。
 *
 * 为什么需要:ExcelJS 只暴露 row.height(= ht 值),**不暴露 customHeight 属性**。而"是否手动设高"
 * 决定渲染层要不要自动行高 —— Excel/WPS 对 customHeight 行从不自动撑高(只裁切/溢出)。少了这个标记,
 * 我们的 autoFitRowHeights 会把"作者手动设矮、但塞了长文本"的行撑大,渲染比 WPS 高、导出也带上多余高度。
 */
import type { RawPackage } from './raw-xml'
import { parseRels, toArray } from './raw-xml'
import type { SheetModel } from '../model/types'

export function attachRowMeta(pkg: RawPackage, sheets: SheetModel[]): void {
  // workbook.xml: sheet name → r:id → sheetN.xml 路径(与 drawing-parser 同套映射)
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
    const xml = pkg.parse(path)
    const rowNodes = toArray(xml?.worksheet?.sheetData?.row)
    for (const rn of rowNodes) {
      const ch = rn['@_customHeight']
      if (ch === 1 || ch === '1' || ch === true) {
        const r0 = Number(rn['@_r']) - 1 // 1-based → 0-based
        if (!Number.isFinite(r0) || r0 < 0) continue
        const cur = sheet.rows.get(r0)
        if (cur) cur.customHeight = true
        // cur 缺失(无 ht 但标了 customHeight,罕见)→ 不强建,默认高已够
      }
    }
  }
}
