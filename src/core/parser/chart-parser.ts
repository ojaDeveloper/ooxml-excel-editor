/**
 * 解析 xl/charts/chartN.xml → ChartSpec(交给 ECharts 映射)。
 * 覆盖常见: bar/col、line、pie/doughnut、area、scatter、radar。
 */
import type { RawPackage } from './raw-xml'
import { toArray } from './raw-xml'
import type { ChartSpec, ChartSeries } from '../model/types'

export function parseChart(pkg: RawPackage, chartPath: string): Omit<ChartSpec, 'anchor'> | undefined {
  const xml = pkg.parse(chartPath)
  const chart = xml?.chartSpace?.chart
  if (!chart) return undefined

  const plotArea = chart.plotArea
  if (!plotArea) return undefined

  const title = extractTitle(chart)
  const showLegend = !!chart.legend

  // 找到第一个已知的图表类型节点
  const typeMap: { key: string; type: ChartSpec['type']; dir?: 'col' | 'bar' }[] = [
    { key: 'barChart', type: 'bar' },
    { key: 'bar3DChart', type: 'bar' },
    { key: 'lineChart', type: 'line' },
    { key: 'line3DChart', type: 'line' },
    { key: 'pieChart', type: 'pie' },
    { key: 'pie3DChart', type: 'pie' },
    { key: 'doughnutChart', type: 'doughnut' },
    { key: 'areaChart', type: 'area' },
    { key: 'area3DChart', type: 'area' },
    { key: 'scatterChart', type: 'scatter' },
    { key: 'radarChart', type: 'radar' },
  ]

  for (const entry of typeMap) {
    const node = plotArea[entry.key]
    if (!node) continue
    const cNode = Array.isArray(node) ? node[0] : node
    let type = entry.type
    let barDirection: 'col' | 'bar' | undefined
    if (entry.key.startsWith('bar')) {
      barDirection = cNode.barDir?.['@_val'] === 'bar' ? 'bar' : 'col'
    }
    const { categories, series } = extractSeries(cNode, type)
    return { type, title, showLegend, barDirection, categories, series }
  }

  return { type: 'unsupported', title, showLegend, categories: [], series: [] }
}

function extractTitle(chart: any): string | undefined {
  const rich = chart.title?.tx?.rich
  if (!rich) return undefined
  const paras = toArray(rich.p)
  const texts: string[] = []
  for (const p of paras) {
    for (const r of toArray(p.r)) {
      if (r.t != null) texts.push(typeof r.t === 'object' ? r.t['#text'] ?? '' : String(r.t))
    }
  }
  return texts.join('') || undefined
}

function extractSeries(chartNode: any, type: ChartSpec['type']): { categories: (string | number)[]; series: ChartSeries[] } {
  const sers = toArray(chartNode.ser)
  const series: ChartSeries[] = []
  let categories: (string | number)[] = []

  sers.forEach((ser: any, i: number) => {
    // 系列名
    let name: string | undefined
    const txStr = ser.tx?.strRef?.strCache ?? ser.tx?.v
    if (ser.tx?.strRef?.strCache) name = readCacheStrings(ser.tx.strRef.strCache)[0]
    else if (typeof txStr === 'string') name = txStr

    // 类别(取第一个系列的 cat)
    if (i === 0) {
      const cat = ser.cat
      if (cat) {
        const strs = cat.strRef?.strCache ? readCacheStrings(cat.strRef.strCache) : null
        const nums = cat.numRef?.numCache ? readCacheNumbers(cat.numRef.numCache) : null
        categories = strs ?? (nums ?? []).map((n) => n ?? '')
      }
    }

    // 值
    let values: (number | null)[] = []
    if (type === 'scatter') {
      const yCache = ser.yVal?.numRef?.numCache
      values = yCache ? readCacheNumbers(yCache) : []
    } else {
      const numCache = ser.val?.numRef?.numCache
      values = numCache ? readCacheNumbers(numCache) : []
    }

    // 颜色(spPr/solidFill)
    const color = readSeriesColor(ser)
    series.push({ name, values, color })
  })

  return { categories, series }
}

function readCacheStrings(cache: any): string[] {
  const pts = toArray(cache.pt)
  const out: string[] = []
  for (const p of pts) {
    const idx = Number(p['@_idx'] ?? out.length)
    out[idx] = p.v != null ? (typeof p.v === 'object' ? p.v['#text'] ?? '' : String(p.v)) : ''
  }
  return Array.from(out, (v) => v ?? '')
}

function readCacheNumbers(cache: any): (number | null)[] {
  const pts = toArray(cache.pt)
  const count = Number(cache['@_count'] ?? pts.length)
  const out: (number | null)[] = new Array(count).fill(null)
  for (const p of pts) {
    const idx = Number(p['@_idx'] ?? 0)
    const v = p.v
    out[idx] = v == null || v === '' ? null : Number(v)
  }
  return out
}

function readSeriesColor(ser: any): string | undefined {
  const fill = ser.spPr?.solidFill
  if (fill?.srgbClr?.['@_val']) return '#' + String(fill.srgbClr['@_val']).toUpperCase()
  return undefined
}
