/**
 * 颜色解析: ExcelJS/OOXML 颜色可能是 argb / theme+tint / indexed。
 * 统一成 css "#RRGGBB" / "rgba(...)"。
 */
import type { CssColor } from '../model/types'

/** Excel 老 indexed 调色板(0-63)。索引 64/65 = 系统前景/背景，调用方特判。 */
const INDEXED_PALETTE: string[] = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
  '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
  '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
]

export function argbToCss(argb: string | undefined): CssColor | undefined {
  if (!argb) return undefined
  let hex = argb.replace(/^#/, '')
  if (hex.length === 8) {
    // AARRGGBB
    const a = parseInt(hex.slice(0, 2), 16) / 255
    const r = parseInt(hex.slice(2, 4), 16)
    const g = parseInt(hex.slice(4, 6), 16)
    const b = parseInt(hex.slice(6, 8), 16)
    if (a >= 1) return `#${hex.slice(2)}`
    return `rgba(${r},${g},${b},${+a.toFixed(3)})`
  }
  if (hex.length === 6) return `#${hex}`
  return undefined
}

/**
 * 任意 css 颜色 → `#RRGGBB`(大写);供 <input type=color> 回显 / 粘贴样式解析。
 * 支持 #RGB / #RRGGBB / #RRGGBBAA / rgb()/rgba();识别不了返 ''。
 */
export function toHex6(css: string | undefined): string {
  if (!css) return ''
  const s = css.trim()
  let m = /^#([0-9a-f]{3})$/i.exec(s) // #RGB → #RRGGBB
  if (m) return ('#' + m[1].split('').map((c) => c + c).join('')).toUpperCase()
  m = /^#([0-9a-f]{6})(?:[0-9a-f]{2})?$/i.exec(s) // #RRGGBB[AA]
  if (m) return '#' + m[1].toUpperCase()
  const rgb = /rgba?\(([^)]+)\)/i.exec(s)
  if (rgb) {
    const p = rgb[1].split(',').map((x) => parseInt(x.trim(), 10))
    if (p.length >= 3 && p.slice(0, 3).every((n) => Number.isFinite(n))) {
      const h = (n: number) => Math.max(0, Math.min(255, n)).toString(16).padStart(2, '0')
      return ('#' + h(p[0]) + h(p[1]) + h(p[2])).toUpperCase()
    }
  }
  return ''
}

export function indexedToCss(idx: number): CssColor | undefined {
  if (idx >= 0 && idx < INDEXED_PALETTE.length) return INDEXED_PALETTE[idx]
  return undefined
}

/** 主题色 + tint → css。tint > 0 变亮，< 0 变暗(按 OOXML 规范的 HSL Luminance 调整)。 */
export function themeToCss(themeColors: CssColor[], themeIdx: number, tint = 0): CssColor | undefined {
  const base = themeColors[themeIdx]
  if (!base) return undefined
  if (!tint) return base
  return applyTint(base, tint)
}

export function applyTint(hexColor: CssColor, tint: number): CssColor {
  const rgb = hexToRgb(hexColor)
  if (!rgb) return hexColor
  const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b)
  // OOXML: lum' = lum*(1+tint) for tint<0 ; lum' = lum*(1-tint)+(1-(1-tint)) ... 采用规范近似
  let lum = hsl.l
  if (tint < 0) {
    lum = lum * (1 + tint)
  } else {
    lum = lum * (1 - tint) + tint
  }
  const out = hslToRgb(hsl.h, hsl.s, clamp01(lum))
  return rgbToHex(out.r, out.g, out.b)
}

/**
 * 统一入口: ExcelJS 风格的颜色对象 → css。
 * 形如 { argb }, { theme, tint }, { indexed }。
 */
export function resolveColor(
  color: any,
  themeColors: CssColor[],
): CssColor | undefined {
  if (!color) return undefined
  if (typeof color === 'string') return argbToCss(color)
  if (typeof color.argb === 'string') return argbToCss(color.argb)
  if (typeof color.theme === 'number') return themeToCss(themeColors, color.theme, color.tint ?? 0)
  if (typeof color.indexed === 'number') return indexedToCss(color.indexed)
  return undefined
}

// ---- helpers ----
function clamp01(v: number) {
  return Math.min(1, Math.max(0, v))
}
function hexToRgb(hex: string): { r: number; g: number; b: number } | null {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex)
  if (!m) return null
  const n = parseInt(m[1], 16)
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 }
}
function rgbToHex(r: number, g: number, b: number): string {
  const h = (v: number) => Math.round(v).toString(16).padStart(2, '0')
  return `#${h(r)}${h(g)}${h(b)}`
}
function rgbToHsl(r: number, g: number, b: number) {
  r /= 255
  g /= 255
  b /= 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  let h = 0
  let s = 0
  const l = (max + min) / 2
  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0)
        break
      case g:
        h = (b - r) / d + 2
        break
      default:
        h = (r - g) / d + 4
    }
    h /= 6
  }
  return { h, s, l }
}
function hslToRgb(h: number, s: number, l: number) {
  let r: number
  let g: number
  let b: number
  if (s === 0) {
    r = g = b = l
  } else {
    const hue2rgb = (p: number, q: number, t: number) => {
      if (t < 0) t += 1
      if (t > 1) t -= 1
      if (t < 1 / 6) return p + (q - p) * 6 * t
      if (t < 1 / 2) return q
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
      return p
    }
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s
    const p = 2 * l - q
    r = hue2rgb(p, q, h + 1 / 3)
    g = hue2rgb(p, q, h)
    b = hue2rgb(p, q, h - 1 / 3)
  }
  return { r: r * 255, g: g * 255, b: b * 255 }
}
