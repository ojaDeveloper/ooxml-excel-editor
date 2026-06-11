import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { parseWorkbook } from '../index'
import { cellKey } from '../../model/types'
import { formatValue } from '../../format/number-format'

function loadSample(): ArrayBuffer {
  const buf = readFileSync(join(__dirname, '..', '..', '..', '..', 'public', 'sample.xlsx'))
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)
}

describe('parseWorkbook (端到端)', () => {
  it('解析示例文件: sheet 数 / 维度 / 主题色', async () => {
    const wb = await parseWorkbook(loadSample())
    expect(wb.sheets.length).toBe(3)
    expect(wb.sheets[0].name).toBe('销售报表')
    expect(wb.themeColors.length).toBeGreaterThanOrEqual(12)
  })

  it('styles[0] 恒为中性空白默认(首格 A1 有绿底也不占 index 0)', async () => {
    // 回归: 解析按遇到顺序 intern 样式, 首格 A1 是绿底表头(#21A366)。
    // 若让 A1 样式占 styles[0], 则所有空格/新建格/setCellValue/applyStyleOverride 兜底(styleId 0)
    // 都会冒出绿底 —— WPS 粘贴 / 编辑串色的根因。styles[0] 必须是无填充中性默认。
    const wb = await parseWorkbook(loadSample())
    const s1 = wb.sheets[0]
    expect(s1.styles[0].fill).toEqual({ type: 'none' })
    expect(s1.styles[0].borders).toEqual({})
    // A1 的绿底样式仍存在(在别的 index 上), 真正引用它的格不受影响
    const a1 = s1.cells.get(cellKey(0, 0))!
    expect(a1.styleId).not.toBe(0)
    expect(s1.styles[a1.styleId].fill).toEqual({ type: 'solid', fgColor: '#21A366' })
  })

  it('合并单元格被识别', async () => {
    const wb = await parseWorkbook(loadSample())
    const s1 = wb.sheets[0]
    expect(s1.merges.length).toBeGreaterThanOrEqual(1)
    // A1:E1 → top0 left0 right4
    expect(s1.merges.some((m) => m.top === 0 && m.left === 0 && m.right === 4)).toBe(true)
  })

  it('冻结窗格 / 自动筛选', async () => {
    const wb = await parseWorkbook(loadSample())
    const s1 = wb.sheets[0]
    expect(s1.freeze.frozenRows).toBe(2)
    expect(s1.freeze.frozenCols).toBe(1)
    expect(s1.autoFilterRange).toBeTruthy()
  })

  it('公式缓存值 + 数字格式正确渲染', async () => {
    const wb = await parseWorkbook(loadSample())
    const s1 = wb.sheets[0]
    // D3 = B3*C3 = 5999*120 = 719880
    const d3 = s1.cells.get(cellKey(2, 3))
    expect(d3).toBeTruthy()
    expect(d3!.type).toBe('formula')
    expect(d3!.raw).toBe(719880)
    const style = s1.styles[d3!.styleId]
    const rendered = formatValue(d3!.raw, style.numFmt, wb.date1904)
    expect(rendered.text).toContain('719,880')
  })

  it('条件格式被解析(dataBar + colorScale)', async () => {
    const wb = await parseWorkbook(loadSample())
    const s1 = wb.sheets[0]
    const types = s1.conditional.map((c) => c.type)
    expect(types).toContain('dataBar')
    expect(types).toContain('colorScale')
  })

  it('大表维度', async () => {
    const wb = await parseWorkbook(loadSample())
    const s3 = wb.sheets[2]
    expect(s3.dimension.rows).toBe(2000)
  })

  it('批注被解析(格式样例 sheet)', async () => {
    const wb = await parseWorkbook(loadSample())
    const s2 = wb.sheets[1]
    const hasComment = [...s2.cells.values()].some((c) => !!c.comment)
    expect(hasComment).toBe(true)
  })

  it('列表数据验证被解析', async () => {
    const wb = await parseWorkbook(loadSample())
    const s2 = wb.sheets[1]
    expect(s2.dataValidations.length).toBeGreaterThanOrEqual(1)
  })

  it('超链接单元格被识别', async () => {
    const wb = await parseWorkbook(loadSample())
    const s2 = wb.sheets[1]
    const hasLink = [...s2.cells.values()].some((c) => !!c.hyperlink)
    expect(hasLink).toBe(true)
  })
})
