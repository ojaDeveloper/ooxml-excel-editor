import { describe, it, expect } from 'vitest'
import { shiftFormulaRefs, type ShiftSpec } from '../refs'

const base = { targetSheet: 'S1', formulaSheet: 'S1' }
const insRow = (at: number, count = 1): ShiftSpec => ({ axis: 'row', at, count, mode: 'insert', ...base })
const delRow = (at: number, count = 1): ShiftSpec => ({ axis: 'row', at, count, mode: 'delete', ...base })
const insCol = (at: number, count = 1): ShiftSpec => ({ axis: 'col', at, count, mode: 'insert', ...base })
const delCol = (at: number, count = 1): ShiftSpec => ({ axis: 'col', at, count, mode: 'delete', ...base })

describe('shiftFormulaRefs 行插入/删除', () => {
  it('插入行:>=at 的行引用下移,之上不动', () => {
    // 在第 5 行(0-based at=4)上方插入 1 行 → A5/B6 下移,A1 不动
    expect(shiftFormulaRefs('=A5+B6+A1', insRow(4))).toBe('=A6+B7+A1')
  })
  it('删除行:被引用行 → #REF!;之后上移', () => {
    expect(shiftFormulaRefs('=A5', delRow(4))).toBe('=#REF!') // 删第 5 行
    expect(shiftFormulaRefs('=A6', delRow(4))).toBe('=A5') // 第 6 行上移
    expect(shiftFormulaRefs('=A1', delRow(4))).toBe('=A1') // 之上不动
  })
  it('绝对引用 $ 保留', () => {
    expect(shiftFormulaRefs('=$A$5+B$6', insRow(4))).toBe('=$A$6+B$7')
  })
})

describe('shiftFormulaRefs 列插入/删除', () => {
  it('插入列:>=at 的列引用右移(字母进位)', () => {
    // 在 B 列(at=1)前插 1 列 → B1→C1,Z1→AA1,A1 不动
    expect(shiftFormulaRefs('=A1+B1+Z1', insCol(1))).toBe('=A1+C1+AA1')
  })
  it('删除列:被引用列 → #REF!', () => {
    expect(shiftFormulaRefs('=B1', delCol(1))).toBe('=#REF!')
    expect(shiftFormulaRefs('=C1', delCol(1))).toBe('=B1')
  })
})

describe('shiftFormulaRefs 区域 / 跨表 / 边界', () => {
  it('区域:删除部分行 → 收缩;全删 → #REF!', () => {
    expect(shiftFormulaRefs('=SUM(A1:A10)', delRow(2, 2))).toBe('=SUM(A1:A8)') // 删第 3-4 行
    expect(shiftFormulaRefs('=SUM(A3:A4)', delRow(2, 2))).toBe('=SUM(#REF!)') // 整段被删
    expect(shiftFormulaRefs('=SUM(A1:A10)', insRow(2))).toBe('=SUM(A1:A11)') // 插入扩展下界
  })
  it('跨表限定:仅指向 targetSheet 的引用重写', () => {
    const spec: ShiftSpec = { axis: 'row', at: 4, count: 1, mode: 'insert', targetSheet: 'S1', formulaSheet: 'S2' }
    expect(shiftFormulaRefs('=S1!A5+A5', spec)).toBe('=S1!A6+A5') // S1!A5 重写;裸 A5 属 S2 不动
    expect(shiftFormulaRefs("='My Sheet'!A5", { ...spec, targetSheet: 'My Sheet' })).toBe("='My Sheet'!A6")
  })
  it('函数名 / 字符串字面量 / 数字不误伤', () => {
    expect(shiftFormulaRefs('=LOG10(A5)', insRow(4))).toBe('=LOG10(A6)') // LOG10 函数不动,A5 重写
    expect(shiftFormulaRefs('=A5&"see A5 text"', insRow(4))).toBe('=A6&"see A5 text"') // 字面量不动
    expect(shiftFormulaRefs('=A1*2', insRow(4))).toBe('=A1*2') // 之上 + 数字常量不动
  })
})
