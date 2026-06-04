/** echarts 按需加载(可选 peer 依赖)。共享给 OverlayManager 与导出图表栅格化。 */
import type * as EChartsNS from 'echarts'

let mod: typeof EChartsNS | null = null

export async function loadECharts(): Promise<typeof EChartsNS> {
  if (!mod) mod = await import('echarts')
  return mod
}
