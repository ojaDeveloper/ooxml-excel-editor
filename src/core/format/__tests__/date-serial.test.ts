/**
 * Excel 日期序列号 ↔ Date 兼容性单测 (2026-06-08).
 *
 * Excel 两个历史包袱必须照搬, 否则日期差 1 天或 4 年:
 *
 *  1. **1900 闰年 bug** — Excel 误把 1900 当闰年(实际不是, 整百年要能被 400 整除).
 *     序号 60 = 不存在的 "1900-02-29". 这是从 Lotus 1-2-3 抄来的 bug, 微软为兼容
 *     **故意保留**. 转换时必须把 ≥ 60 的序号 -1 天, 否则 1900-03-01 之后所有日期差 1 天.
 *
 *  2. **1904 系统**(Mac Excel) — workbookPr@date1904="1". 序号 0 = 1904-01-01, 无闰年 bug.
 *     同一序号在两套系统下含义差 **1462 天** (~4 年 1 天).
 *
 * 测试目标: 锁住 1900 边界 + 1904 全套 + 双向往返一致.
 */
import { describe, it, expect } from 'vitest'
import { serialToDate, serialTimePart } from '../date-serial'
import { dateToSerial } from '../number-format'

/** 把 Date 转为 YYYY-MM-DD (UTC) 便于断言 */
function ymd(d: Date): string {
  return [
    d.getUTCFullYear(),
    String(d.getUTCMonth() + 1).padStart(2, '0'),
    String(d.getUTCDate()).padStart(2, '0'),
  ].join('-')
}

describe('serialToDate — 1900 系统 (默认 date1904=false)', () => {
  it('序号 1 → 1900-01-01 (1900 系统起点)', () => {
    expect(ymd(serialToDate(1))).toBe('1900-01-01')
  })

  it('序号 59 → 1900-02-28 (闰年 bug 修正前的最后一天)', () => {
    expect(ymd(serialToDate(59))).toBe('1900-02-28')
  })

  it('序号 60 → 1900-02-28 (phantom 1900-02-29; 修正后跟序号 59 同日 — Excel 真实行为)', () => {
    // 序号 60 是 Excel 的"不存在的 2 月 29 日". serialToDate 走 s>=60 → s-=1 分支 → 落回 1900-02-28.
    // 这跟 Excel 实际显示 "1900-02-29" 不一致, 但跟"正确的真实日期"一致.
    // 跟 LibreOffice / Google Sheets 行为相同 (它们都把 60 当 1900-02-28 显示).
    expect(ymd(serialToDate(60))).toBe('1900-02-28')
  })

  it('序号 61 → 1900-03-01 (闰年 bug 修正后, 跳过 phantom 日)', () => {
    expect(ymd(serialToDate(61))).toBe('1900-03-01')
  })

  it('1900-03-01 之后日期都受 -1 修正; 2000-01-01 = 序号 36526', () => {
    expect(ymd(serialToDate(36526))).toBe('2000-01-01')
  })

  it('远期: 2023-12-25 = 序号 45285', () => {
    expect(ymd(serialToDate(45285))).toBe('2023-12-25')
  })

  it('带小数: 序号 45285.5 = 2023-12-25 12:00 (中午)', () => {
    const d = serialToDate(45285.5)
    expect(ymd(d)).toBe('2023-12-25')
    expect(d.getUTCHours()).toBe(12)
    expect(d.getUTCMinutes()).toBe(0)
  })
})

describe('serialToDate — 1904 系统 (Mac Excel; date1904=true)', () => {
  it('序号 0 → 1904-01-01 (1904 系统起点)', () => {
    expect(ymd(serialToDate(0, true))).toBe('1904-01-01')
  })

  it('序号 1 → 1904-01-02', () => {
    expect(ymd(serialToDate(1, true))).toBe('1904-01-02')
  })

  it('无 phantom 闰年: 序号 59 → 1904-02-29 (1904 是真闰年)', () => {
    // 1904 真闰年 (能被 400 整除). 序号 0~58 = 1904-01-01..1904-02-28, 序号 59 = 1904-02-29
    expect(ymd(serialToDate(59, true))).toBe('1904-02-29')
  })

  it('序号 60 → 1904-03-01 (1904 系统下 60 是真实日期, 不是 phantom)', () => {
    expect(ymd(serialToDate(60, true))).toBe('1904-03-01')
  })

  it('远期: 2023-12-25 在 1904 系统下 = 序号 43823 (比 1900 系统少 1462 天)', () => {
    // 同一物理日期, 序号差 = (1904-01-01 - 1899-12-30) - 1(因为 1904 epoch 是 day 0 不是 day 1) = 1462
    expect(ymd(serialToDate(43823, true))).toBe('2023-12-25')
  })

  it('同一序号在两套系统下相差 4 年 1 天 (1462 天)', () => {
    // 序号 45285 在 1900 = 2023-12-25, 在 1904 = ~2027-12-26
    const d1900 = serialToDate(45285, false)
    const d1904 = serialToDate(45285, true)
    const diffDays = (d1904.getTime() - d1900.getTime()) / 86400000
    expect(diffDays).toBe(1462)
  })
})

describe('dateToSerial — 反向转换', () => {
  it('1900 系统: 2023-12-25 → 序号 45285', () => {
    expect(dateToSerial(new Date(Date.UTC(2023, 11, 25)), false)).toBe(45285)
  })

  it('1900 系统: 1900-01-01 → 序号 1', () => {
    expect(dateToSerial(new Date(Date.UTC(1900, 0, 1)), false)).toBe(1)
  })

  it('1900 系统: 1900-03-01 → 序号 61 (跳过 phantom 60)', () => {
    expect(dateToSerial(new Date(Date.UTC(1900, 2, 1)), false)).toBe(61)
  })

  it('1904 系统: 2023-12-25 → 序号 43823', () => {
    expect(dateToSerial(new Date(Date.UTC(2023, 11, 25)), true)).toBe(43823)
  })

  it('1904 系统: 1904-01-01 → 序号 0', () => {
    expect(dateToSerial(new Date(Date.UTC(1904, 0, 1)), true)).toBe(0)
  })
})

describe('往返一致性 — serialToDate ∘ dateToSerial', () => {
  it('1900 系统: 序号 → Date → 序号 在 phantom 之外保持不变', () => {
    for (const n of [1, 59, 61, 100, 36526, 45285]) {
      const round = dateToSerial(serialToDate(n, false), false)
      expect(round).toBe(n)
    }
  })

  it('1904 系统: 序号 → Date → 序号 完全保持', () => {
    for (const n of [0, 1, 59, 60, 100, 43823, 50000]) {
      const round = dateToSerial(serialToDate(n, true), true)
      expect(round).toBe(n)
    }
  })

  it('1900 系统序号 60 (phantom) 往返折回 59 (跟 LibreOffice/Google Sheets 一致)', () => {
    // serialToDate(60) = 1900-02-28 (跟 59 同日); dateToSerial(1900-02-28) = 59
    const round = dateToSerial(serialToDate(60, false), false)
    expect(round).toBe(59) // 不是 60, 因为 phantom 日不存在
  })
})

describe('serialTimePart — 一天内的时分秒', () => {
  it('整数部分: 时分秒全 0', () => {
    expect(serialTimePart(100)).toEqual({ h: 0, m: 0, s: 0, frac: 0 })
  })

  it('0.5 = 12:00', () => {
    expect(serialTimePart(100.5)).toMatchObject({ h: 12, m: 0, s: 0 })
  })

  it('0.25 = 06:00', () => {
    expect(serialTimePart(0.25)).toMatchObject({ h: 6, m: 0, s: 0 })
  })

  it('0.75 = 18:00', () => {
    expect(serialTimePart(0.75)).toMatchObject({ h: 18, m: 0, s: 0 })
  })
})
