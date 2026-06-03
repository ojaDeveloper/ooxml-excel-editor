import { describe, expect, it } from 'vitest'
import { parsePageSetup } from '../exceljs-adapter'

// parsePageSetup 只读 ws.pageSetup,用最小 mock 即可(无需真实 ExcelJS 实例)
const mk = (ps: any) => parsePageSetup({ pageSetup: ps } as any)

describe('parsePageSetup', () => {
  it('无 pageSetup → undefined', () => {
    expect(parsePageSetup({} as any)).toBeUndefined()
  })

  it('方向/缩放/适应页面', () => {
    const r = mk({ orientation: 'landscape', scale: 80, fitToPage: true, fitToWidth: 1, fitToHeight: 0 })
    expect(r?.orientation).toBe('landscape')
    expect(r?.scale).toBe(80)
    expect(r?.fitToPage).toBe(true)
    expect(r?.fitToWidth).toBe(1)
  })

  it('paperSize 代码映射(9→a4, 1→letter, 8→a3, 未知→省略)', () => {
    expect(mk({ paperSize: 9 })?.paperFormat).toBe('a4')
    expect(mk({ paperSize: 1 })?.paperFormat).toBe('letter')
    expect(mk({ paperSize: 8 })?.paperFormat).toBe('a3')
    expect(mk({ paperSize: 999 })?.paperFormat).toBeUndefined()
  })

  it('页边距 inch → mm', () => {
    const r = mk({ margins: { left: 1, right: 1, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 } })
    expect(r?.margins?.left).toBeCloseTo(25.4, 4)
    expect(r?.margins?.top).toBeCloseTo(12.7, 4)
    expect(r?.margins?.header).toBeCloseTo(7.62, 4)
  })

  it('printArea 取首段 + 去 Sheet! 前缀 → 0-based 闭区间', () => {
    const r = mk({ printArea: 'Sheet1!$A$1:$C$5,Sheet1!$E$1:$F$2' })
    expect(r?.printArea).toEqual({ top: 0, left: 0, bottom: 4, right: 2 })
  })

  it('打印标题行 "1:2" → [0,1];标题列 "A:B" → [0,1]', () => {
    const r = mk({ printTitlesRow: '1:2', printTitlesColumn: 'A:B' })
    expect(r?.printTitleRows).toEqual([0, 1])
    expect(r?.printTitleCols).toEqual([0, 1])
  })

  it('带 $ 与 Sheet! 前缀的标题行也能解析', () => {
    const r = mk({ printTitlesRow: 'Sheet1!$3:$5' })
    expect(r?.printTitleRows).toEqual([2, 4])
  })
})
