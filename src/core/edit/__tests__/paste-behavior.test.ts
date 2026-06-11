import { describe, it, expect } from 'vitest'
import { makeDefaultStyle, type CellStyle, type CellStyleOverride } from '../../model/types'
import {
  DEFAULT_PASTE_BEHAVIOR,
  PASTE_PRESET_VALUES_ONLY,
  resolvePasteBehavior,
  shouldApplyColWidth,
  resolvePastedCellStyle,
} from '../paste-behavior'

const neutral = (): CellStyle => makeDefaultStyle()
const red = (): CellStyle => ({ ...makeDefaultStyle(), fill: { type: 'solid', fgColor: '#FF0000' }, hAlign: 'right', font: { ...makeDefaultStyle().font, bold: true } })

describe('PasteBehavior 默认 / 预设 / 补全', () => {
  it('默认 = 覆盖式 1:1;列宽 firstRowOnly;目标合并 clear', () => {
    expect(DEFAULT_PASTE_BEHAVIOR.cellStyle).toBe('overwrite')
    expect(DEFAULT_PASTE_BEHAVIOR.fill).toBe('overwrite')
    expect(DEFAULT_PASTE_BEHAVIOR.colWidth).toBe('firstRowOnly')
    expect(DEFAULT_PASTE_BEHAVIOR.targetMerges).toBe('clear')
  })
  it('仅值预设 = 啥都不动只填值', () => {
    expect(PASTE_PRESET_VALUES_ONLY).toMatchObject({ cellStyle: 'skip', fill: 'skip', rowHeight: 'keep', colWidth: 'keep', sourceMerges: 'skip', targetMerges: 'keep', images: 'skip' })
  })
  it('resolvePasteBehavior 缺项回落默认', () => {
    expect(resolvePasteBehavior({ cellStyle: 'merge' })).toMatchObject({ cellStyle: 'merge', fill: 'overwrite', colWidth: 'firstRowOnly' })
    expect(resolvePasteBehavior(null)).toEqual(DEFAULT_PASTE_BEHAVIOR)
  })
  it('shouldApplyColWidth: source 总搬 / firstRowOnly 仅 row0 / keep 不搬', () => {
    expect(shouldApplyColWidth({ ...DEFAULT_PASTE_BEHAVIOR, colWidth: 'source' }, 5)).toBe(true)
    expect(shouldApplyColWidth({ ...DEFAULT_PASTE_BEHAVIOR, colWidth: 'firstRowOnly' }, 0)).toBe(true)
    expect(shouldApplyColWidth({ ...DEFAULT_PASTE_BEHAVIOR, colWidth: 'firstRowOnly' }, 5)).toBe(false)
    expect(shouldApplyColWidth({ ...DEFAULT_PASTE_BEHAVIOR, colWidth: 'keep' }, 0)).toBe(false)
  })
})

describe('resolvePastedCellStyle — 样式/填充各档', () => {
  // 源 patch:有边框、居中,但没写填充
  const patch: CellStyleOverride = { hAlign: 'center', borders: { top: { style: 'thin', color: '#000000' } } }

  it('overwrite/overwrite:目标红底被清成无填充(贴近源),非填充取源', () => {
    const r = resolvePastedCellStyle(red(), neutral(), patch, 'overwrite', 'overwrite')!
    expect(r.fill).toEqual({ type: 'none' }) // 源没写填充 → 无填充(不漏红底)
    expect(r.hAlign).toBe('center')
    expect(r.font.bold).toBe(false) // 取中性默认(目标的 bold 被丢)
    expect(r.borders.top?.style).toBe('thin')
  })

  it('merge/merge:源没写的保留目标(红底 + 目标 bold 都留)', () => {
    const r = resolvePastedCellStyle(red(), neutral(), patch, 'merge', 'merge')!
    expect(r.fill).toEqual({ type: 'solid', fgColor: '#FF0000' }) // 源没写填充 → 留目标红底
    expect(r.hAlign).toBe('center') // 源写了 → 覆盖
    expect(r.font.bold).toBe(true) // 源没写 → 留目标 bold
  })

  it('skip/skip:样式整个不动 → 返 null', () => {
    expect(resolvePastedCellStyle(red(), neutral(), patch, 'skip', 'skip')).toBeNull()
  })

  it('混档 overwrite 样式 + skip 填充:非填充取源,填充留目标', () => {
    const r = resolvePastedCellStyle(red(), neutral(), patch, 'overwrite', 'skip')!
    expect(r.hAlign).toBe('center')
    expect(r.fill).toEqual({ type: 'solid', fgColor: '#FF0000' }) // fill skip → 留目标
  })

  it('源写了填充:overwrite/merge 都取源填充', () => {
    const withFill: CellStyleOverride = { fill: { type: 'solid', fgColor: '#00FF00' } }
    expect(resolvePastedCellStyle(red(), neutral(), withFill, 'overwrite', 'overwrite')!.fill).toEqual({ type: 'solid', fgColor: '#00FF00' })
    expect(resolvePastedCellStyle(red(), neutral(), withFill, 'merge', 'merge')!.fill).toEqual({ type: 'solid', fgColor: '#00FF00' })
  })
})
