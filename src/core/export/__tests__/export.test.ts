import { describe, expect, it } from 'vitest'
import { MM_PER_PX, resolveMargins, resolvePageSize } from '../raster'
import { exportToPdf } from '../pdf'
import { exportToVectorPdf, hasNonLatin, hexToRgb, packBands } from '../vector-pdf'
import { makeBands } from '../paginate'

// 注: 测试环境为 node(无 DOM/canvas),只覆盖纯函数与边界;
// 渲染/切片/合成依赖真实 canvas,放到浏览器/手动验证。

describe('resolvePageSize', () => {
  it('a4 纵向 = 210x297', () => {
    expect(resolvePageSize('a4', 'portrait')).toEqual([210, 297])
  })
  it('a4 横向 = 297x210(宽高互换)', () => {
    expect(resolvePageSize('a4', 'landscape')).toEqual([297, 210])
  })
  it('letter / a3', () => {
    expect(resolvePageSize('letter', 'portrait')).toEqual([215.9, 279.4])
    expect(resolvePageSize('a3', 'landscape')).toEqual([420, 297])
  })
  it('自定义 [宽,高] 数组,纵向取 min 宽 max 高', () => {
    expect(resolvePageSize([100, 200], 'portrait')).toEqual([100, 200])
    expect(resolvePageSize([200, 100], 'portrait')).toEqual([100, 200])
    expect(resolvePageSize([100, 200], 'landscape')).toEqual([200, 100])
  })
  it('缺省 = a4 纵向', () => {
    expect(resolvePageSize()).toEqual([210, 297])
  })
})

describe('resolveMargins', () => {
  it('数字 → 四边相同', () => {
    expect(resolveMargins(12)).toEqual({ top: 12, right: 12, bottom: 12, left: 12 })
  })
  it('部分对象 → 缺省补 10', () => {
    expect(resolveMargins({ top: 5, left: 8 })).toEqual({ top: 5, right: 10, bottom: 10, left: 8 })
  })
  it('undefined → 全 10', () => {
    expect(resolveMargins(undefined)).toEqual({ top: 10, right: 10, bottom: 10, left: 10 })
  })
})

describe('MM_PER_PX', () => {
  it('按 96dpi: 96px = 25.4mm', () => {
    expect(MM_PER_PX * 96).toBeCloseTo(25.4, 6)
  })
})

describe('exportToPdf 边界', () => {
  it('空工作表列表 → 抛友好错误', async () => {
    await expect(exportToPdf([])).rejects.toThrow('没有可导出的工作表')
  })
})

describe('矢量 PDF 工具', () => {
  it('exportToVectorPdf 空列表 → 抛友好错误', async () => {
    await expect(exportToVectorPdf([])).rejects.toThrow('没有可导出的工作表')
  })

  it('hasNonLatin: 纯拉丁/数字 false,含中文 true', () => {
    expect(hasNonLatin('Total 123.45')).toBe(false)
    expect(hasNonLatin('日期')).toBe(true)
    expect(hasNonLatin('ABC 你好')).toBe(true)
    expect(hasNonLatin('')).toBe(false)
    expect(hasNonLatin('café résumé')).toBe(false) // 拉丁扩展 ≤0x250
  })

  it('hexToRgb: 解析 #RRGGBB,非法回黑', () => {
    expect(hexToRgb('#FF8000')).toEqual([255, 128, 0])
    expect(hexToRgb('00FF00')).toEqual([0, 255, 0])
    expect(hexToRgb('bad')).toEqual([0, 0, 0])
  })
})

describe('分页 packBands / makeBands(横向分页核心)', () => {
  it('packBands: 等宽列按上限分带,每带至少一项', () => {
    // 每项 50,上限 120 → 每带最多 2 项(100≤120,150>120)
    expect(packBands([0, 1, 2, 3, 4], () => 50, 120)).toEqual([[0, 1], [2, 3], [4]])
    // 单项超上限也独占一带(不丢)
    expect(packBands([0, 1], () => 200, 120)).toEqual([[0], [1]])
    // 空 → 一个空带
    expect(packBands([], () => 50, 120)).toEqual([[]])
  })

  it('makeBands: 像素按带宽切分,长度累计=总量', () => {
    const bands = makeBands(1000, 300)
    expect(bands.map((b) => b.len).reduce((a, b) => a + b, 0)).toBe(1000)
    expect(bands.length).toBe(4) // 300+300+300+100
    expect(bands[0]).toEqual({ s: 0, len: 300 })
    expect(bands[3]).toEqual({ s: 900, len: 100 })
  })

  it('makeBands: 带宽≥总量 → 单带', () => {
    expect(makeBands(500, 800)).toEqual([{ s: 0, len: 500 }])
  })
})
