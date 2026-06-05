/**
 * 公式引擎抽象(框架 + 库无关)—— core 只依赖这个接口,默认实现是 HyperFormula 适配器,
 * 用方可注入自研/其它引擎(EditConfig.formulaEngine)。这是"可换引擎"的承重接口。
 *
 * 地址:全 0-based(sheet 索引 / row / col),与模型一致。公式文本(cell.formula)是 A1 串,
 * 引擎自己解析 A1,故这层不做 A1↔0-based 转换 —— 转换风险只在公式串内部,由引擎承担。
 */
import type { CellModel, WorkbookModel } from '../model/types'
import type { CellValue } from '../model/data-access'

/** 一个被引擎判定为"值变了"的格(含算出的新值)。sheet = 0-based 工作表索引。 */
export interface DirtyCell {
  sheet: number
  row: number
  col: number
  value: CellValue
}

export interface FormulaEngine {
  /** 用整个工作簿(全表,供跨表引用)初始化引擎。 */
  setSheets(wb: WorkbookModel): void
  /** 设一个格的内容(公式串 '=...' 或字面值或 null),返回所有受影响格(含依赖级联)。 */
  setCell(sheet: number, row: number, col: number, content: string | number | boolean | null): DirtyCell[]
  /** 读一个格的当前计算值。 */
  getValue(sheet: number, row: number, col: number): CellValue
  /** 释放引擎资源。 */
  destroy(): void
}

/** 引擎工厂:异步(默认适配器懒 import hyperformula);用方可注入。 */
export type FormulaEngineFactory = () => Promise<FormulaEngine>

/**
 * 取一个格喂给引擎的内容:公式格 → 公式文本(A1 串);值格 → 字面值;空/日期 → 见下。
 * 日期暂以 null 入引擎(v1 不支持公式引用日期格运算,避免类型不匹配;后续可换 Excel 序列号)。
 */
export function cellContentForEngine(cell: CellModel | null): string | number | boolean | null {
  if (!cell || cell.type === 'empty') return null
  if (cell.type === 'formula') {
    const f = cell.formula
    if (f == null || f === '') return null
    // 解析层存的公式不带前导 '='(ExcelJS v.formula='B3*C3'),用户输入的带('=A1+1')→ 统一补 '=' 喂引擎
    return f[0] === '=' ? f : '=' + f
  }
  const raw = cell.raw
  if (typeof raw === 'number' || typeof raw === 'boolean' || typeof raw === 'string') return raw
  return null // Date / 其它:v1 不入引擎
}

export type { CellValue }
