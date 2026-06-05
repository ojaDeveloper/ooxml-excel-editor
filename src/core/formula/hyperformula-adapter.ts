/**
 * 默认公式引擎适配器 —— HyperFormula。仅 `import type`(类型,构建后擦除),运行时靠工厂里的
 * 动态 `import('hyperformula')` 懒加载,故 hyperformula 是**可选 peer**、不进 core 产物。
 *
 * 许可证:HyperFormula 是 GPL-3.0 / 商业 双授权。这里用 `licenseKey: 'gpl-v3'`(开源场景)。
 * 商业项目请用方注入自己持牌的引擎(EditConfig.formulaEngine),或换其它引擎。详见 README。
 */
import type { HyperFormula, DetailedCellError } from 'hyperformula'
import type { WorkbookModel } from '../model/types'
import type { CellValue } from '../model/data-access'
import { cellContentForEngine, type DirtyCell, type FormulaEngine, type FormulaEngineFactory } from './engine'

type HFStatic = typeof HyperFormula
type DetailedCellErrorCtor = typeof DetailedCellError

class HyperFormulaAdapter implements FormulaEngine {
  private hf: HyperFormula | null = null
  private sheetIds: number[] = [] // modelIndex → hf sheetId(按名解析,避免数字名重排坑)

  constructor(
    private HF: HFStatic,
    private DCE: DetailedCellErrorCtor,
  ) {}

  setSheets(wb: WorkbookModel): void {
    const sheetsObj: Record<string, (string | number | boolean | null)[][]> = {}
    const names: string[] = []
    for (let i = 0; i < wb.sheets.length; i++) {
      const s = wb.sheets[i]
      let name = s.name || `Sheet${i + 1}`
      while (names.includes(name)) name += '_' // 去重(无效 xlsx 防御)
      names.push(name)
      const rows = Math.max(s.dimension.rows, 1)
      const cols = Math.max(s.dimension.cols, 1)
      const grid: (string | number | boolean | null)[][] = []
      for (let r = 0; r < rows; r++) grid.push(new Array(cols).fill(null))
      for (const cell of s.cells.values()) {
        if (cell.row < rows && cell.col < cols) grid[cell.row][cell.col] = cellContentForEngine(cell)
      }
      sheetsObj[name] = grid
    }
    this.hf = this.HF.buildFromSheets(sheetsObj, { licenseKey: 'gpl-v3' })
    this.sheetIds = names.map((n) => this.hf!.getSheetId(n) ?? 0)
  }

  setCell(sheet: number, row: number, col: number, content: string | number | boolean | null): DirtyCell[] {
    if (!this.hf) return []
    const id = this.sheetIds[sheet] ?? sheet
    const changes = this.hf.setCellContents({ sheet: id, row, col }, content)
    const idToIndex = (hfId: number) => {
      const idx = this.sheetIds.indexOf(hfId)
      return idx >= 0 ? idx : hfId
    }
    const out: DirtyCell[] = []
    for (const ch of changes) {
      if (!('address' in ch)) continue // 跳过命名表达式变更(无格地址)
      out.push({
        sheet: idToIndex(ch.address.sheet),
        row: ch.address.row,
        col: ch.address.col,
        value: this.mapValue(ch.newValue),
      })
    }
    return out
  }

  getValue(sheet: number, row: number, col: number): CellValue {
    if (!this.hf) return null
    const id = this.sheetIds[sheet] ?? sheet
    return this.mapValue(this.hf.getCellValue({ sheet: id, row, col }))
  }

  destroy(): void {
    this.hf?.destroy()
    this.hf = null
  }

  /** HF 值 → 模型 CellValue:错误对象 → 错误串(如 #DIV/0!);空 → null。 */
  private mapValue(v: unknown): CellValue {
    if (v instanceof this.DCE) return v.value
    if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v
    return null
  }
}

/** 默认引擎工厂:动态 import hyperformula(不进 core 产物)。装了才用,没装则 import 抛错由调用方兜底。 */
export const defaultFormulaEngineFactory: FormulaEngineFactory = async () => {
  const mod = await import('hyperformula')
  const HF = (mod.HyperFormula ?? mod.default) as HFStatic
  const DCE = mod.DetailedCellError as DetailedCellErrorCtor
  return new HyperFormulaAdapter(HF, DCE)
}
