/** 合并区域索引: 快速判断某 cell 是否在合并区、是否为锚点、跨度多少。 */
import type { MergeRange, SheetModel } from '../model/types'

export class MergeIndex {
  /** key `${row}:${col}` → 该 cell 所属的合并区(任意成员都能查到) */
  private map = new Map<string, MergeRange>()
  /** 锚点(左上角) key 集合 */
  private anchors = new Set<string>()

  constructor(sheet: SheetModel) {
    for (const m of sheet.merges) {
      this.anchors.add(`${m.top}:${m.left}`)
      for (let r = m.top; r <= m.bottom; r++) {
        for (let c = m.left; c <= m.right; c++) {
          this.map.set(`${r}:${c}`, m)
        }
      }
    }
  }

  /** 该 cell 所在的合并区(没有返回 undefined) */
  rangeOf(row: number, col: number): MergeRange | undefined {
    return this.map.get(`${row}:${col}`)
  }
  /** 是否合并区锚点(左上角，渲染内容的那个) */
  isAnchor(row: number, col: number): boolean {
    return this.anchors.has(`${row}:${col}`)
  }
  /** 是否被合并覆盖但不是锚点(渲染时跳过内容) */
  isCovered(row: number, col: number): boolean {
    const m = this.map.get(`${row}:${col}`)
    return !!m && !(m.top === row && m.left === col)
  }
}
