/** ChartSpec → ECharts option(近似还原)。 */
import type { ChartSpec } from '../model/types'
import type { EChartsOption } from 'echarts'

export function chartToOption(spec: ChartSpec): EChartsOption {
  const base: EChartsOption = {
    animation: false,
    title: spec.title ? { text: spec.title, left: 'center', textStyle: { fontSize: 13 } } : undefined,
    legend: spec.showLegend ? { bottom: 0, type: 'scroll' } : undefined,
    grid: { left: 40, right: 16, top: spec.title ? 36 : 16, bottom: spec.showLegend ? 32 : 24, containLabel: true },
    tooltip: { trigger: spec.type === 'pie' || spec.type === 'doughnut' ? 'item' : 'axis' },
  }

  const palette = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E']

  if (spec.type === 'pie' || spec.type === 'doughnut') {
    const s = spec.series[0]
    const data = (s?.values ?? []).map((v, i) => ({ name: String(spec.categories[i] ?? i + 1), value: v ?? 0 }))
    return {
      ...base,
      tooltip: { trigger: 'item' },
      series: [
        {
          type: 'pie',
          radius: spec.type === 'doughnut' ? ['40%', '70%'] : '70%',
          data,
          label: { fontSize: 10 },
        },
      ],
      color: palette,
    }
  }

  if (spec.type === 'scatter') {
    return {
      ...base,
      xAxis: { type: 'value', scale: true },
      yAxis: { type: 'value', scale: true },
      series: spec.series.map((s, i) => ({
        name: s.name,
        type: 'scatter',
        data: s.values.map((v, idx) => [Number(spec.categories[idx] ?? idx), v]),
        itemStyle: { color: s.color || palette[i % palette.length] },
      })),
    }
  }

  if (spec.type === 'radar') {
    const indicators = spec.categories.map((c) => ({ name: String(c) }))
    return {
      ...base,
      radar: { indicator: indicators.length ? indicators : [{ name: '' }] },
      series: [
        {
          type: 'radar',
          data: spec.series.map((s, i) => ({
            name: s.name,
            value: s.values.map((v) => v ?? 0),
            itemStyle: { color: s.color || palette[i % palette.length] },
          })),
        },
      ],
    }
  }

  // bar / line / area
  const categoryAxis = { type: 'category' as const, data: spec.categories.map(String), axisLabel: { fontSize: 10 } }
  const valueAxis = { type: 'value' as const, axisLabel: { fontSize: 10 } }
  const horizontal = spec.type === 'bar' && spec.barDirection === 'bar'

  return {
    ...base,
    xAxis: horizontal ? valueAxis : categoryAxis,
    yAxis: horizontal ? categoryAxis : valueAxis,
    series: spec.series.map((s, i) => ({
      name: s.name,
      type: spec.type === 'line' || spec.type === 'area' ? 'line' : 'bar',
      areaStyle: spec.type === 'area' ? {} : undefined,
      data: s.values.map((v) => v ?? null),
      itemStyle: { color: s.color || palette[i % palette.length] },
      smooth: false,
    })),
    color: palette,
  }
}
