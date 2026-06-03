/**
 * 解析每个 worksheet 关联的 xl/drawings/drawingN.xml，抽出图片与图表锚点，
 * 回填到对应 SheetModel.images / charts。
 *
 * 链路: workbook.xml(顺序+r:id) → workbook.xml.rels(sheet 路径)
 *      → sheetN.xml.rels(drawing 路径) → drawingN.xml(锚点 + pic/graphicFrame)
 *      → drawingN.xml.rels(图片/图表真实路径)
 */
import type { RawPackage } from './raw-xml'
import { parseRels, toArray, basename } from './raw-xml'
import { parseChart } from './chart-parser'
import type { AnchorCell, ChartSpec, ImageAnchor, SheetModel } from '../model/types'

const MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
}

export function attachDrawings(pkg: RawPackage, sheets: SheetModel[]): void {
  // 1. workbook.xml: sheet name → r:id(顺序)
  const wbXml = pkg.parse('xl/workbook.xml')
  const sheetNodes = toArray(wbXml?.workbook?.sheets?.sheet)
  const wbRels = parseRels(pkg, 'xl/workbook.xml')

  const nameToPath = new Map<string, string>()
  for (const sn of sheetNodes) {
    const name = sn['@_name']
    const rid = sn['@_id'] // r:id → id(removeNSPrefix)
    if (name && rid && wbRels[rid]) nameToPath.set(String(name), wbRels[rid])
  }

  for (const sheet of sheets) {
    const sheetPath = nameToPath.get(sheet.name)
    if (!sheetPath) continue
    const sheetRels = parseRels(pkg, sheetPath)
    // 找 drawing 关系
    const sheetXml = pkg.parse(sheetPath)
    const drawingRef = sheetXml?.worksheet?.drawing
    const drawingRid = Array.isArray(drawingRef) ? drawingRef[0]?.['@_id'] : drawingRef?.['@_id']
    if (!drawingRid || !sheetRels[drawingRid]) continue
    const drawingPath = sheetRels[drawingRid]
    parseDrawing(pkg, drawingPath, sheet)
  }
}

function parseDrawing(pkg: RawPackage, drawingPath: string, sheet: SheetModel): void {
  const xml = pkg.parse(drawingPath)
  const wsDr = xml?.wsDr
  if (!wsDr) return
  const drawingRels = parseRels(pkg, drawingPath)

  const anchors = [
    ...toArray(wsDr.twoCellAnchor).map((a: any) => ({ a, kind: 'two' as const })),
    ...toArray(wsDr.oneCellAnchor).map((a: any) => ({ a, kind: 'one' as const })),
  ]

  for (const { a, kind } of anchors) {
    const from = readAnchorCell(a.from)
    const to = kind === 'two' ? readAnchorCell(a.to) ?? undefined : undefined
    let extW: number | undefined
    let extH: number | undefined
    if (kind === 'one' && a.ext) {
      extW = Number(a.ext['@_cx'])
      extH = Number(a.ext['@_cy'])
    }
    if (!from) continue
    const anchorBase: ImageAnchor = { src: '', from, to, extWidthEmu: extW, extHeightEmu: extH }

    // 图片
    const pic = a.pic
    if (pic) {
      const embed = pic.blipFill?.blip?.['@_embed'] ?? pic.blipFill?.blip?.['@_link']
      const target = embed ? drawingRels[embed] : undefined
      if (target) {
        const data = imageData(pkg, target)
        if (data) sheet.images.push({ ...anchorBase, src: '', bytes: data.bytes, mime: data.mime })
      }
      continue
    }

    // 图表(graphicFrame → c:chart r:id)
    const gf = a.graphicFrame
    if (gf) {
      const chartRef = gf.graphic?.graphicData?.chart
      const chartRid = chartRef?.['@_id']
      const chartPath = chartRid ? drawingRels[chartRid] : undefined
      if (chartPath) {
        const spec = parseChart(pkg, chartPath)
        if (spec && spec.type !== 'unsupported') {
          const chart: ChartSpec = { ...spec, anchor: anchorBase }
          sheet.charts.push(chart)
        }
      }
    }
  }
}

function readAnchorCell(node: any): AnchorCell | null {
  if (!node) return null
  return {
    col: Number(node.col ?? 0),
    colOffEmu: Number(node.colOff ?? 0),
    row: Number(node.row ?? 0),
    rowOffEmu: Number(node.rowOff ?? 0),
  }
}

function imageData(pkg: RawPackage, path: string): { bytes: Uint8Array; mime: string } | undefined {
  const bytes = pkg.bytes(path)
  if (!bytes) return undefined
  const ext = basename(path).split('.').pop()?.toLowerCase() || 'png'
  const mime = MIME[ext]
  if (!mime) return undefined // emf/wmf 等浏览器不支持，跳过
  return { bytes, mime }
}
