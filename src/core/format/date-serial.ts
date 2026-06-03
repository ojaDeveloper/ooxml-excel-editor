/**
 * Excel 日期序列号 ↔ JS Date。
 *
 * 1900 系统: 序列号 1 = 1900-01-01。但 Excel 误把 1900 当闰年(序列号 60 = 不存在的
 * 1900-02-29)，这是为兼容 Lotus 1-2-3 故意保留的 bug。所以 >= 60 的序列号要减 1 天。
 * 1904 系统(老 Mac): 序列号 0 = 1904-01-01，无闰年 bug。
 */

const MS_PER_DAY = 86400 * 1000

// 1899-12-31 00:00:00 UTC 作为 1900 系统的基准(序列号 1 → 1900-01-01)
const EPOCH_1900 = Date.UTC(1899, 11, 31)
// 1904-01-01
const EPOCH_1904 = Date.UTC(1904, 0, 1)

/** 序列号(可能带小数表示时间) → Date(UTC 基准，渲染按本地读字段) */
export function serialToDate(serial: number, date1904 = false): Date {
  if (date1904) {
    return new Date(EPOCH_1904 + serial * MS_PER_DAY)
  }
  // 1900 系统: 修正 1900-02-29 幽灵日
  let s = serial
  if (s >= 60) s -= 1
  return new Date(EPOCH_1900 + s * MS_PER_DAY)
}

/** 取序列号的小数部分对应的一天内秒数(用于纯时间格式) */
export function serialTimePart(serial: number): { h: number; m: number; s: number; frac: number } {
  const frac = serial - Math.floor(serial)
  let totalSec = Math.round(frac * 86400)
  const h = Math.floor(totalSec / 3600)
  totalSec -= h * 3600
  const m = Math.floor(totalSec / 60)
  const s = totalSec - m * 60
  return { h, m, s, frac }
}
