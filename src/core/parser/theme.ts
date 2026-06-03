/**
 * 解析 xl/theme/theme1.xml 的配色方案 → 主题色数组。
 * 顺序按 ECMA-376 clrScheme: dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink。
 *
 * 注意 Excel 渲染时主题索引顺序与 clrScheme 声明顺序不同(dk1/lt1 互换):
 * theme 索引 0=lt1(背景1) 1=dk1(文本1) 2=lt2 3=dk2 4..9=accent1..6 10=hlink 11=folHlink。
 * 这里产出"按 theme 索引"对齐的数组，供 color.themeToCss 直接用。
 */
import type { RawPackage } from './raw-xml'
import { toArray } from './raw-xml'
import type { CssColor } from '../model/types'

const DEFAULT_THEME: CssColor[] = [
  '#FFFFFF', // 0 lt1 (bg1)
  '#000000', // 1 dk1 (tx1)
  '#E7E6E6', // 2 lt2 (bg2)
  '#44546A', // 3 dk2 (tx2)
  '#4472C4', // 4 accent1
  '#ED7D31', // 5 accent2
  '#A5A5A5', // 6 accent3
  '#FFC000', // 7 accent4
  '#5B9BD5', // 8 accent5
  '#70AD47', // 9 accent6
  '#0563C1', // 10 hlink
  '#954F72', // 11 folHlink
]

export function parseTheme(pkg: RawPackage): CssColor[] {
  try {
    const xml = pkg.parse('xl/theme/theme1.xml')
    const scheme = xml?.theme?.themeElements?.clrScheme
    if (!scheme) return [...DEFAULT_THEME]

    const read = (node: any): CssColor | undefined => {
      if (!node) return undefined
      if (node.srgbClr) return '#' + String(node.srgbClr['@_val']).toUpperCase()
      if (node.sysClr) {
        const last = node.sysClr['@_lastClr']
        if (last) return '#' + String(last).toUpperCase()
      }
      return undefined
    }

    const dk1 = read(scheme.dk1)
    const lt1 = read(scheme.lt1)
    const dk2 = read(scheme.dk2)
    const lt2 = read(scheme.lt2)
    const accents = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map(
      (k) => read(scheme[k]),
    )
    const hlink = read(scheme.hlink)
    const folHlink = read(scheme.folHlink)

    // 按 theme 索引顺序: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink
    const ordered = [lt1, dk1, lt2, dk2, ...accents, hlink, folHlink]
    return ordered.map((c, i) => c ?? DEFAULT_THEME[i] ?? '#000000')
  } catch {
    return [...DEFAULT_THEME]
  }
  // toArray imported for potential multi-scheme handling; keep referenced.
  void toArray
}
