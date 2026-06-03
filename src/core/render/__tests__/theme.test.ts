import { describe, it, expect } from 'vitest'
import { DEFAULT_THEME, mergeTheme } from '../theme'

describe('mergeTheme', () => {
  it('无参数返回默认主题副本', () => {
    expect(mergeTheme()).toEqual(DEFAULT_THEME)
    expect(mergeTheme()).not.toBe(DEFAULT_THEME) // 是副本,不可变默认
  })

  it('部分覆盖,其余保留默认', () => {
    const t = mergeTheme({ gridLine: '#ff0000', selBorder: '#000' })
    expect(t.gridLine).toBe('#ff0000')
    expect(t.selBorder).toBe('#000')
    expect(t.headerBg).toBe(DEFAULT_THEME.headerBg) // 未覆盖项保持默认
  })
})
