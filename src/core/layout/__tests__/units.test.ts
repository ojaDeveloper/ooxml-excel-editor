import { describe, it, expect } from 'vitest'
import { colWidthToPx, rowHeightToPx } from '../units'

describe('列宽换算 colWidthToPx', () => {
  it('默认列宽 8.43 字符 = 64px(Excel 实测值)', () => {
    expect(colWidthToPx(8.43)).toBe(64)
  })
  it('未指定(null)回退到默认 = 64px', () => {
    expect(colWidthToPx(null as unknown as number)).toBe(64)
  })
  it('0 宽(隐藏列)= 0', () => {
    expect(colWidthToPx(0)).toBe(0)
  })
  it('含内边距: 每列比裸 width*7 多 5px', () => {
    // 旧公式 floor(10*7)=70 偏窄; 正确 = floor(70.5)+5 = 75
    expect(colWidthToPx(10)).toBe(75)
  })
})

describe('行高换算 rowHeightToPx', () => {
  it('默认行高 15pt = 20px', () => {
    expect(rowHeightToPx(15)).toBe(20)
  })
})
