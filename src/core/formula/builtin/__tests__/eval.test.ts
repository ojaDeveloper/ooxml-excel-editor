import { describe, it, expect } from 'vitest'
import { parseFormula } from '../parse'
import { evalAst, type RefResolver } from '../eval'
import { isMatrix, isErr, type Scalar } from '../values'

/** 用一组 A1 地址→值 的表跑公式;无 resolver 时引用都为 null。 */
function run(formula: string, cells: Record<string, Scalar> = {}): Scalar | string {
  const res: RefResolver = {
    getCell: (_sheet, row, col) => {
      const a1 = String.fromCharCode(65 + col) + (row + 1)
      return cells[a1] ?? null
    },
  }
  const v = evalAst(parseFormula(formula), res)
  const s = isMatrix(v) ? v.rows[0][0] : v
  return isErr(s) ? s.code : s
}

describe('内置公式 求值 — 运算符', () => {
  it('算术 + 优先级 + 括号', () => {
    expect(run('=1+2*3')).toBe(7)
    expect(run('=(1+2)*3')).toBe(9)
    expect(run('=2^10')).toBe(1024)
    expect(run('=7/2')).toBe(3.5)
    expect(run('=10/0')).toBe('#DIV/0!')
    expect(run('=-3+5')).toBe(2)
    expect(run('=50%')).toBe(0.5)
  })
  it('连接 & + 比较', () => {
    expect(run('="a"&"b"&1')).toBe('ab1')
    expect(run('=1<2')).toBe(true)
    expect(run('=2<>2')).toBe(false)
    expect(run('="A"="a"')).toBe(true) // 文本比较不区分大小写
  })
  it('引用参与运算', () => {
    expect(run('=A1+B1', { A1: 3, B1: 4 })).toBe(7)
    expect(run('=A1&B1', { A1: 'x', B1: 2 })).toBe('x2')
  })
})

describe('内置公式 求值 — 函数', () => {
  it('聚合 SUM/AVERAGE/MAX/MIN/COUNT', () => {
    const c = { A1: 1, A2: 2, A3: 3, A4: 'x' as Scalar }
    expect(run('=SUM(A1:A4)', c)).toBe(6)
    expect(run('=AVERAGE(A1:A3)', c)).toBe(2)
    expect(run('=MAX(A1:A4)', c)).toBe(3)
    expect(run('=MIN(A1:A4)', c)).toBe(1)
    expect(run('=COUNT(A1:A4)', c)).toBe(3) // 文本不计
    expect(run('=COUNTA(A1:A4)', c)).toBe(4)
  })
  it('IF / IFERROR / AND / OR / NOT', () => {
    expect(run('=IF(A1>0,"pos","neg")', { A1: 5 })).toBe('pos')
    expect(run('=IF(A1>0,"pos","neg")', { A1: -5 })).toBe('neg')
    expect(run('=IFERROR(1/0, 99)')).toBe(99)
    expect(run('=AND(TRUE, 1>0)')).toBe(true)
    expect(run('=OR(FALSE, 0)')).toBe(false)
    expect(run('=NOT(TRUE)')).toBe(false)
  })
  it('文本 LEFT/RIGHT/MID/LEN/CONCAT/UPPER', () => {
    expect(run('=LEFT("hello",2)')).toBe('he')
    expect(run('=RIGHT("hello",2)')).toBe('lo')
    expect(run('=MID("hello",2,3)')).toBe('ell')
    expect(run('=LEN("hello")')).toBe(5)
    expect(run('=CONCAT("a","b","c")')).toBe('abc')
    expect(run('=UPPER("abc")')).toBe('ABC')
  })
  it('数学 ROUND/ABS/INT/MOD/SQRT/POWER', () => {
    expect(run('=ROUND(3.14159,2)')).toBe(3.14)
    expect(run('=ABS(-7)')).toBe(7)
    expect(run('=INT(3.9)')).toBe(3)
    expect(run('=MOD(7,3)')).toBe(1)
    expect(run('=SQRT(16)')).toBe(4)
    expect(run('=POWER(2,8)')).toBe(256)
  })
  it('条件聚合 SUMIF/COUNTIF', () => {
    const c = { A1: 5, A2: 15, A3: 25, B1: 1, B2: 2, B3: 3 }
    expect(run('=SUMIF(A1:A3,">10")', c)).toBe(40)
    expect(run('=COUNTIF(A1:A3,">10")', c)).toBe(2)
    expect(run('=SUMIF(A1:A3,">10",B1:B3)', c)).toBe(5) // 对应 B2+B3
  })
  it('查找 VLOOKUP/INDEX/MATCH', () => {
    const c = { A1: 'a', B1: 10, A2: 'b', B2: 20, A3: 'c', B3: 30 }
    expect(run('=VLOOKUP("b",A1:B3,2,FALSE)', c)).toBe(20)
    expect(run('=INDEX(A1:B3,3,2)', c)).toBe(30)
    expect(run('=MATCH("c",A1:A3,0)', c)).toBe(3)
  })
  it('错误传播 + #NAME? 未知函数', () => {
    expect(run('=SUM(1, 1/0)')).toBe('#DIV/0!')
    expect(run('=FOOBAR(1)')).toBe('#NAME?')
    expect(run('=ISERROR(1/0)')).toBe(true)
  })
  it('日期 DATE/YEAR/MONTH/DAY', () => {
    const serial = run('=DATE(2026,4,1)') as number
    expect(typeof serial).toBe('number')
    expect(run(`=YEAR(${serial})`)).toBe(2026)
    expect(run(`=MONTH(${serial})`)).toBe(4)
    expect(run(`=DAY(${serial})`)).toBe(1)
  })
})
