/** OOXML 单位换算常量与函数。 */

/** 默认字体(Calibri 11)的最大数字宽度，单位 px。Excel 列宽换算的基准。 */
export const DEFAULT_MDW = 7
/** 屏幕 DPI / 印刷 DPI */
export const PX_PER_POINT = 96 / 72 // 1pt = 4/3 px
/** 1 px = 9525 EMU (914400 EMU/inch ÷ 96 px/inch) */
export const EMU_PER_PX = 9525

export const DEFAULT_COL_WIDTH_CHARS = 8.43
export const DEFAULT_ROW_HEIGHT_PT = 15

/**
 * Excel 列宽(字符数) → 像素。
 * px = Truncate(width * MDW + 0.5) + 5
 * 末尾 +5 是单元格内边距(左右各 2px + 1px 网格线)。这是 Excel 实际渲染宽度,
 * 早期只用 Truncate(width*MDW) 会每列偏窄 5px(默认 8.43 → 64px 才对)。
 */
export function colWidthToPx(chars: number, mdw = DEFAULT_MDW): number {
  if (chars == null) chars = DEFAULT_COL_WIDTH_CHARS
  if (chars <= 0) return 0
  return Math.floor(chars * mdw + 0.5) + 5
}

/** 行高(pt) → 像素 */
export function rowHeightToPx(pt: number): number {
  return Math.round((pt ?? DEFAULT_ROW_HEIGHT_PT) * PX_PER_POINT)
}

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX
}
