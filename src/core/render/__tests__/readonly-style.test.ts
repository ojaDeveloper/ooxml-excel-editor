/**
 * Phase C: cellStyle 钩子 ctx.editable + readOnlyCellStyle 三态 单测 (2026-06-08).
 *
 * 验证 CanvasRenderer.styleOf 的钩子链:
 *   ① cellStyle 收到第 3 入参 ctx.editable
 *   ② readOnlyCellStyle=true 时只读格套内置浅灰 fill
 *   ③ readOnlyCellStyle 为 CellStyleOverride 对象时套固定样式
 *   ④ readOnlyCellStyle 为 CellStyleFn 时按格调用
 *   ⑤ editable=true 时 readOnlyCellStyle 不生效
 */
import { describe, it, expect } from 'vitest'
import type { CellStyleFn, CellStyleOverride } from '../../model/types'
import { cellKey } from '../../model/types'
import { jsonToWorkbook } from '../../loader-json'

/** mock canvas 2d context 给 CanvasRenderer 构造用 */
function makeMockCanvas(): HTMLCanvasElement {
  const fakeCtx = {
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 1,
    font: '',
    textBaseline: 'middle',
    textAlign: 'left',
    canvas: { width: 800, height: 600 },
    measureText: () => ({ width: 50 }),
    fillRect: () => {},
    strokeRect: () => {},
    fillText: () => {},
    strokeText: () => {},
    save: () => {},
    restore: () => {},
    beginPath: () => {},
    moveTo: () => {},
    lineTo: () => {},
    rect: () => {},
    arc: () => {},
    fill: () => {},
    stroke: () => {},
    clip: () => {},
    closePath: () => {},
    translate: () => {},
    rotate: () => {},
    scale: () => {},
    setTransform: () => {},
    drawImage: () => {},
    setLineDash: () => {},
  }
  const canvas = {
    getContext: () => fakeCtx,
    width: 800,
    height: 600,
    style: {},
  } as unknown as HTMLCanvasElement
  return canvas
}

describe('CanvasRenderer.styleOf — Phase C 只读视觉钩子 (2026-06-08)', () => {
  it('cellStyle 钩子收到第 3 入参 ctx.editable', async () => {
    const { CanvasRenderer } = await import('../canvas-renderer')
    const wb = jsonToWorkbook([['a', 'b'], ['c', 'd']])
    const sheet = wb.sheets[0]
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'a', styleId: 0 })

    const seenCtx: Array<unknown> = []
    const hook: CellStyleFn = (cell, _pos, ctx) => {
      seenCtx.push({ row: cell.row, col: cell.col, editable: ctx?.editable })
      return undefined
    }
    const renderer = new CanvasRenderer(makeMockCanvas(), sheet, wb, 1, {
      cellStyle: hook,
      // isEditable: (0,0) → false, 其它 → true
      isEditable: (r, c) => !(r === 0 && c === 0),
    })
    // 触发 styleOf —— 通过私有访问 (TS 类型上越界, 在测试里 ok)
    const styleOf = (renderer as any).styleOf.bind(renderer)
    styleOf(sheet.cells.get(cellKey(0, 0)))
    expect(seenCtx).toHaveLength(1)
    expect(seenCtx[0]).toEqual({ row: 0, col: 0, editable: false })
  })

  it('readOnlyCellStyle=true: 只读格套内置浅灰 fill #f5f7fa', async () => {
    const { CanvasRenderer } = await import('../canvas-renderer')
    const wb = jsonToWorkbook([['a']])
    const sheet = wb.sheets[0]
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'a', styleId: 0 })

    const renderer = new CanvasRenderer(makeMockCanvas(), sheet, wb, 1, {
      readOnlyCellStyle: true,
      isEditable: () => false, // 该格只读
    })
    const styleOf = (renderer as any).styleOf.bind(renderer)
    const s = styleOf(sheet.cells.get(cellKey(0, 0)))
    expect(s.fill.type).toBe('solid')
    expect(s.fill.fgColor).toBe('#f5f7fa')
  })

  it('readOnlyCellStyle 为 CellStyleOverride: 套固定样式给所有只读格', async () => {
    const { CanvasRenderer } = await import('../canvas-renderer')
    const wb = jsonToWorkbook([['a']])
    const sheet = wb.sheets[0]
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'a', styleId: 0 })

    const customRoStyle: CellStyleOverride = {
      fill: { type: 'solid', fgColor: '#ffeb99' }, // 黄底
      font: { color: '#806600', italic: true } as any,
    }
    const renderer = new CanvasRenderer(makeMockCanvas(), sheet, wb, 1, {
      readOnlyCellStyle: customRoStyle,
      isEditable: () => false,
    })
    const styleOf = (renderer as any).styleOf.bind(renderer)
    const s = styleOf(sheet.cells.get(cellKey(0, 0)))
    expect(s.fill.fgColor).toBe('#ffeb99')
    expect(s.font.color).toBe('#806600')
    expect(s.font.italic).toBe(true)
  })

  it('readOnlyCellStyle 为 CellStyleFn: 按格调用, 函数自定义返样式', async () => {
    const { CanvasRenderer } = await import('../canvas-renderer')
    const wb = jsonToWorkbook([['a']])
    const sheet = wb.sheets[0]
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'a', styleId: 0 })
    sheet.cells.set(cellKey(0, 1), { row: 0, col: 1, type: 'string', raw: 'b', styleId: 0 })

    const calls: Array<{ row: number; col: number; editable: boolean }> = []
    const roFn: CellStyleFn = (cell, _pos, ctx) => {
      calls.push({ row: cell.row, col: cell.col, editable: !!ctx?.editable })
      return cell.col === 1 ? { fill: { type: 'solid', fgColor: '#ff0000' } } : undefined
    }
    const renderer = new CanvasRenderer(makeMockCanvas(), sheet, wb, 1, {
      readOnlyCellStyle: roFn,
      isEditable: () => false, // 两格都只读 → 都过 readOnlyCellStyle
    })
    const styleOf = (renderer as any).styleOf.bind(renderer)
    const s0 = styleOf(sheet.cells.get(cellKey(0, 0)))
    const s1 = styleOf(sheet.cells.get(cellKey(0, 1)))
    expect(calls).toHaveLength(2)
    expect(calls[0].editable).toBe(false) // ctx.editable 透传
    // (0,0) 函数返 undefined → 不动 fill
    expect(s0.fill.type).not.toBe('solid')
    // (0,1) 函数返红色 → 套上
    expect(s1.fill.fgColor).toBe('#ff0000')
  })

  it('editable=true 时 readOnlyCellStyle 不生效 (只对只读格套用)', async () => {
    const { CanvasRenderer } = await import('../canvas-renderer')
    const wb = jsonToWorkbook([['a']])
    const sheet = wb.sheets[0]
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'a', styleId: 0 })

    const renderer = new CanvasRenderer(makeMockCanvas(), sheet, wb, 1, {
      readOnlyCellStyle: true, // 设了
      isEditable: () => true, // 但全可编辑 → 不套
    })
    const styleOf = (renderer as any).styleOf.bind(renderer)
    const s = styleOf(sheet.cells.get(cellKey(0, 0)))
    expect(s.fill.type).not.toBe('solid')
  })

  it('readOnlyCellStyle=undefined (默认): 老行为, 无视觉差异', async () => {
    const { CanvasRenderer } = await import('../canvas-renderer')
    const wb = jsonToWorkbook([['a']])
    const sheet = wb.sheets[0]
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'a', styleId: 0 })

    const renderer = new CanvasRenderer(makeMockCanvas(), sheet, wb, 1, {
      // 不传 readOnlyCellStyle
      isEditable: () => false, // 即便只读
    })
    const styleOf = (renderer as any).styleOf.bind(renderer)
    const s = styleOf(sheet.cells.get(cellKey(0, 0)))
    // 默认 jsonToWorkbook 给的样式 fill.type = 'none'
    expect(s.fill.type).toBe('none')
  })

  it('isEditable 不注入 (默认): 全格 editable=true, readOnlyCellStyle 套不上 (老行为)', async () => {
    const { CanvasRenderer } = await import('../canvas-renderer')
    const wb = jsonToWorkbook([['a']])
    const sheet = wb.sheets[0]
    sheet.cells.set(cellKey(0, 0), { row: 0, col: 0, type: 'string', raw: 'a', styleId: 0 })

    const renderer = new CanvasRenderer(makeMockCanvas(), sheet, wb, 1, {
      readOnlyCellStyle: true,
      // 不传 isEditable → 默认 () => true → 该格 editable → 不套 RO
    })
    const styleOf = (renderer as any).styleOf.bind(renderer)
    const s = styleOf(sheet.cells.get(cellKey(0, 0)))
    expect(s.fill.type).not.toBe('solid')
  })
})
