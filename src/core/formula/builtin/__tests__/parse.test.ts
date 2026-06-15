import { describe, it, expect } from 'vitest'
import { parseFormula, colToIndex, type Node } from '../parse'

describe('内置公式解析 parseFormula', () => {
  it('colToIndex: A=0, Z=25, AA=26, AB=27', () => {
    expect(colToIndex('A')).toBe(0)
    expect(colToIndex('Z')).toBe(25)
    expect(colToIndex('AA')).toBe(26)
    expect(colToIndex('AB')).toBe(27)
  })
  it('数字 / 字符串 / 布尔', () => {
    expect(parseFormula('=1.5')).toEqual({ t: 'num', v: 1.5 })
    expect(parseFormula('="hi"')).toEqual({ t: 'str', v: 'hi' })
    expect(parseFormula('=TRUE')).toEqual({ t: 'bool', v: true })
  })
  it('引用:相对/绝对/混合 + 0-based 行列', () => {
    expect(parseFormula('=A1')).toMatchObject({ t: 'ref', row: 0, col: 0, absRow: false, absCol: false, sheet: null })
    expect(parseFormula('=$B$2')).toMatchObject({ t: 'ref', row: 1, col: 1, absRow: true, absCol: true })
    expect(parseFormula('=$C4')).toMatchObject({ t: 'ref', row: 3, col: 2, absRow: false, absCol: true })
  })
  it('区域 A1:B3', () => {
    const n = parseFormula('=A1:B3') as Extract<Node, { t: 'range' }>
    expect(n.t).toBe('range')
    expect(n.a).toMatchObject({ row: 0, col: 0 })
    expect(n.b).toMatchObject({ row: 2, col: 1 })
  })
  it('跨表引用 Sheet2!A1 / 带引号', () => {
    expect(parseFormula('=Sheet2!A1')).toMatchObject({ t: 'ref', sheet: 'Sheet2', row: 0, col: 0 })
    expect(parseFormula("='My Sheet'!B2")).toMatchObject({ t: 'ref', sheet: 'My Sheet', row: 1, col: 1 })
  })
  it('运算符优先级:1+2*3 → 1+(2*3)', () => {
    expect(parseFormula('=1+2*3')).toEqual({ t: 'bin', op: '+', l: { t: 'num', v: 1 }, r: { t: 'bin', op: '*', l: { t: 'num', v: 2 }, r: { t: 'num', v: 3 } } })
  })
  it('乘方右结合:2^3^2 → 2^(3^2)', () => {
    const n = parseFormula('=2^3^2') as Extract<Node, { t: 'bin' }>
    expect(n.op).toBe('^')
    expect((n.r as Extract<Node, { t: 'bin' }>).op).toBe('^')
  })
  it('一元负 + 百分号:-5% ', () => {
    expect(parseFormula('=-5%')).toEqual({ t: 'unary', op: '-', x: { t: 'unary', op: '%', x: { t: 'num', v: 5 } } })
  })
  it('函数调用 + 嵌套 + 区域参数', () => {
    const n = parseFormula('=SUM(A1:A3, IF(B1>0, 1, 2))') as Extract<Node, { t: 'func' }>
    expect(n.name).toBe('SUM')
    expect(n.args.length).toBe(2)
    expect(n.args[0].t).toBe('range')
    expect((n.args[1] as Extract<Node, { t: 'func' }>).name).toBe('IF')
  })
  it('比较 + 连接', () => {
    expect((parseFormula('=A1&"x"') as Extract<Node, { t: 'bin' }>).op).toBe('&')
    expect((parseFormula('=A1<>B1') as Extract<Node, { t: 'bin' }>).op).toBe('<>')
  })
  it('括号改变优先级:(1+2)*3', () => {
    const n = parseFormula('=(1+2)*3') as Extract<Node, { t: 'bin' }>
    expect(n.op).toBe('*')
    expect((n.l as Extract<Node, { t: 'bin' }>).op).toBe('+')
  })
  it('错误字面量 #N/A', () => {
    expect(parseFormula('=#N/A')).toEqual({ t: 'err', v: '#N/A' })
  })
})
