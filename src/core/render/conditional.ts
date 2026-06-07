/**
 * 条件格式引擎: 预计算各规则在其范围内的统计量(min/mid/max)，
 * 提供 effectsFor(row,col) 给渲染器叠加 背景色/字体色/数据条/图标。
 * 覆盖 cellIs / colorScale / dataBar / iconSet / top10。
 */
import type { ConditionalRule, MergeRange, SheetModel } from '../model/types'
import { cellKey } from '../model/types'

export interface CellEffect {
  fillColor?: string
  fontColor?: string
  bold?: boolean
  dataBar?: { ratio: number; color: string; gradient: boolean }
  icon?: { setName: string; level: number; count: number }
}

interface PreparedRule {
  rule: ConditionalRule
  min: number
  max: number
  mid: number
  /** top10 阈值 */
  threshold?: number
  inRange(row: number, col: number): boolean
}

export class ConditionalEngine {
  private prepared: PreparedRule[] = []

  constructor(private sheet: SheetModel) {
    for (const rule of sheet.conditional) {
      this.prepared.push(this.prepare(rule))
    }
  }

  hasRules(): boolean {
    return this.prepared.length > 0
  }

  private numericValuesIn(ranges: MergeRange[]): number[] {
    const vals: number[] = []
    for (const rg of ranges) {
      for (let r = rg.top; r <= rg.bottom; r++) {
        for (let c = rg.left; c <= rg.right; c++) {
          const cell = this.sheet.cells.get(cellKey(r, c))
          if (cell && typeof cell.raw === 'number') vals.push(cell.raw)
        }
      }
    }
    return vals
  }

  private prepare(rule: ConditionalRule): PreparedRule {
    const inRange = (row: number, col: number) =>
      rule.ranges.some((rg) => row >= rg.top && row <= rg.bottom && col >= rg.left && col <= rg.right)

    let min = 0
    let max = 0
    let mid = 0
    let threshold: number | undefined

    if (rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'iconSet' || rule.type === 'top10') {
      const vals = this.numericValuesIn(rule.ranges)
      if (vals.length) {
        min = Math.min(...vals)
        max = Math.max(...vals)
        mid = min + (max - min) / 2
        if (rule.type === 'top10') {
          const n = Math.max(1, Math.round(Number(rule.formulae?.[0] ?? 10)))
          const sorted = [...vals].sort((a, b) => b - a)
          threshold = sorted[Math.min(n, sorted.length) - 1]
        }
      }
    }
    return { rule, min, max, mid, threshold, inRange }
  }

  /** 返回所有命中该格的规则索引 + 各自计算出的 effect(供 Cell Inspector 查询;不做"第一条赢"短路) */
  inspectHits(
    row: number,
    col: number,
    value: number | string | boolean | Date | null,
  ): Array<{ ruleIndex: number; effect: CellEffect }> {
    const hits: Array<{ ruleIndex: number; effect: CellEffect }> = []
    for (let i = 0; i < this.prepared.length; i++) {
      const p = this.prepared[i]
      if (!p.inRange(row, col)) continue
      const effect = this.evalRule(p, value)
      if (effect) hits.push({ ruleIndex: i, effect })
    }
    return hits
  }

  /** 单条规则在某格上的 effect(命中返 patch,不命中返 null);effectsFor 与 inspectHits 共用 */
  private evalRule(p: PreparedRule, value: number | string | boolean | Date | null): CellEffect | null {
    const rule = p.rule
    const num = typeof value === 'number' ? value : null
    switch (rule.type) {
      case 'cellIs':
      case 'expression':
      case 'top10': {
        const hit =
          rule.type === 'top10'
            ? num != null && p.threshold != null && num >= p.threshold
            : evalCellIs(rule, num)
        if (!hit || !rule.style) return null
        const patch: CellEffect = {}
        if (rule.style.fill?.fgColor) patch.fillColor = rule.style.fill.fgColor
        if (rule.style.font?.color) patch.fontColor = rule.style.font.color
        if (rule.style.font?.bold) patch.bold = true
        return Object.keys(patch).length ? patch : null
      }
      case 'colorScale': {
        if (num == null || !rule.colorScale || p.max === p.min) return null
        return { fillColor: colorScaleColor(rule.colorScale, num, p.min, p.mid, p.max) }
      }
      case 'dataBar': {
        if (num == null || !rule.dataBar) return null
        const ratio = p.max === p.min ? 1 : Math.max(0, Math.min(1, (num - Math.min(0, p.min)) / (p.max - Math.min(0, p.min))))
        return { dataBar: { ratio, color: rule.dataBar.color, gradient: rule.dataBar.gradient } }
      }
      case 'iconSet': {
        if (num == null || !rule.iconSet) return null
        const count = iconCount(rule.iconSet.name)
        const pos = p.max === p.min ? 1 : (num - p.min) / (p.max - p.min)
        const level = Math.min(count - 1, Math.floor(pos * count))
        return { icon: { setName: rule.iconSet.name, level, count } }
      }
    }
    return null
  }

  effectsFor(row: number, col: number, value: number | string | boolean | Date | null): CellEffect | null {
    let effect: CellEffect | null = null
    const set = (patch: CellEffect) => {
      effect = effect ? { ...patch, ...effect } : patch // 先匹配(高优先级)的 win
    }

    for (const p of this.prepared) {
      if (!p.inRange(row, col)) continue
      const rule = p.rule
      const num = typeof value === 'number' ? value : null

      switch (rule.type) {
        case 'cellIs':
        case 'expression':
        case 'top10': {
          const hit =
            rule.type === 'top10'
              ? num != null && p.threshold != null && num >= p.threshold
              : evalCellIs(rule, num)
          if (hit && rule.style) {
            const patch: CellEffect = {}
            if (rule.style.fill?.fgColor) patch.fillColor = rule.style.fill.fgColor
            if (rule.style.font?.color) patch.fontColor = rule.style.font.color
            if (rule.style.font?.bold) patch.bold = true
            set(patch)
          }
          break
        }
        case 'colorScale': {
          if (num == null || !rule.colorScale || p.max === p.min) break
          set({ fillColor: colorScaleColor(rule.colorScale, num, p.min, p.mid, p.max) })
          break
        }
        case 'dataBar': {
          if (num == null || !rule.dataBar) break
          const ratio = p.max === p.min ? 1 : Math.max(0, Math.min(1, (num - Math.min(0, p.min)) / (p.max - Math.min(0, p.min))))
          set({ dataBar: { ratio, color: rule.dataBar.color, gradient: rule.dataBar.gradient } })
          break
        }
        case 'iconSet': {
          if (num == null || !rule.iconSet) break
          const count = iconCount(rule.iconSet.name)
          const pos = p.max === p.min ? 1 : (num - p.min) / (p.max - p.min)
          const level = Math.min(count - 1, Math.floor(pos * count))
          set({ icon: { setName: rule.iconSet.name, level, count } })
          break
        }
      }
    }
    return effect
  }
}

function evalCellIs(rule: ConditionalRule, num: number | null): boolean {
  if (num == null || !rule.formulae) return false
  const a = Number(rule.formulae[0])
  const b = rule.formulae[1] != null ? Number(rule.formulae[1]) : undefined
  switch (rule.operator) {
    case 'greaterThan': return num > a
    case 'greaterThanOrEqual': return num >= a
    case 'lessThan': return num < a
    case 'lessThanOrEqual': return num <= a
    case 'equal': return num === a
    case 'notEqual': return num !== a
    case 'between': return b != null && num >= a && num <= b
    case 'notBetween': return b != null && (num < a || num > b)
    default: return false
  }
}

function colorScaleColor(
  cs: { min: string; mid?: string; max: string },
  value: number,
  min: number,
  mid: number,
  max: number,
): string {
  if (cs.mid) {
    if (value <= mid) return lerpColor(cs.min, cs.mid, (value - min) / (mid - min || 1))
    return lerpColor(cs.mid, cs.max, (value - mid) / (max - mid || 1))
  }
  return lerpColor(cs.min, cs.max, (value - min) / (max - min || 1))
}

function lerpColor(a: string, b: string, t: number): string {
  t = Math.max(0, Math.min(1, t))
  const ca = hex(a)
  const cb = hex(b)
  const r = Math.round(ca.r + (cb.r - ca.r) * t)
  const g = Math.round(ca.g + (cb.g - ca.g) * t)
  const bl = Math.round(ca.b + (cb.b - ca.b) * t)
  return `rgb(${r},${g},${bl})`
}
function hex(c: string): { r: number; g: number; b: number } {
  const m = /^#?([0-9a-f]{6})$/i.exec(c)
  if (!m) return { r: 255, g: 255, b: 255 }
  const n = parseInt(m[1], 16)
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 }
}

function iconCount(name: string): number {
  if (name.startsWith('5')) return 5
  if (name.startsWith('4')) return 4
  return 3
}
