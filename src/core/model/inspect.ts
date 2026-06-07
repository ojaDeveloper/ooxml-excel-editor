/**
 * Cell Inspector —— 单元格"全息体检"。把分散在 model / view / render 里的查询统一聚合:
 *   - 现有 CellSnapshot 全量(cell + style + raw/computed/text)
 *   - 该格所属合并区(锚点 / 被覆盖)
 *   - 覆盖到该格的浮动图(可多)
 *   - WPS 单元格内嵌图(DISPIMG)
 *   - 命中的数据验证范围(模型层只存范围;详细规则字段当前未建模)
 *   - 命中的条件格式规则(可多;走 ConditionalEngine.inspectHits)
 *   - hyperlink / comment 直读
 *
 * 框架无关、不依赖 renderer;只读 sheet + workbook,可被 headless 流程(导出/批处理)安全调用。
 */
import type { CellStyleOverride, ImageAnchor, MergeRange, SheetModel, WorkbookModel } from './types'
import { cellKey } from './types'
import type { CellSnapshot } from './snapshot'
import { buildCellSnapshot } from './snapshot'
import { ConditionalEngine } from '../render/conditional'

export interface CellInspection extends CellSnapshot {
  /** 该格所属合并区(锚点格 = 自身;被覆盖格 = 覆盖它的;否则 null) */
  merge: MergeRange | null
  /** 是否合并区的左上锚点 */
  isMergeAnchor: boolean
  /** 覆盖到该格的浮动图(按 sheet.images 索引 + 锚点克隆) */
  floatingImages: Array<{ index: number; anchor: ImageAnchor }>
  /** WPS 单元格内嵌图(DISPIMG) */
  cellImage: { id: string; src: string; mime?: string } | null
  /** 命中的数据验证范围(无则 null;详细规则字段当前模型未建模) */
  dataValidation: MergeRange | null
  /** 命中的条件格式规则索引 + 该规则计算后的等效样式 */
  conditional: Array<{ ruleIndex: number; style: CellStyleOverride }>
  /** 直读 cell.hyperlink(便于不深入 cell 字段) */
  hyperlink: string | null
  /** 直读 cell.comment */
  comment: string | null
}

/** 该格落在哪个合并区(锚点 / 被覆盖均算);否则 null */
export function findMergeAt(sheet: SheetModel, row: number, col: number): MergeRange | null {
  for (const m of sheet.merges) {
    if (row >= m.top && row <= m.bottom && col >= m.left && col <= m.right) return m
  }
  return null
}

/** 一张浮动图的锚点是否覆盖到 (row,col)。twoCellAnchor 用 from..to 矩形;oneCellAnchor 仅 from 格。 */
export function imageAnchorContains(anchor: ImageAnchor, row: number, col: number): boolean {
  const f = anchor.from
  if (!f) return false
  if (anchor.to) {
    const t = anchor.to
    return row >= f.row && row <= t.row && col >= f.col && col <= t.col
  }
  return row === f.row && col === f.col
}

/**
 * 聚合查询。
 * @param sheet 工作表
 * @param workbook 工作簿(用于 WPS 内嵌图登记表 cellImages)
 * @param row 0-based
 * @param col 0-based
 * @param date1904 仅作 buildCellSnapshot 文本格式化用
 */
export function inspectCell(
  sheet: SheetModel,
  workbook: WorkbookModel,
  row: number,
  col: number,
  date1904: boolean,
): CellInspection {
  const snapshot = buildCellSnapshot(sheet, row, col, date1904)
  const live = sheet.cells.get(cellKey(row, col))

  const merge = findMergeAt(sheet, row, col)
  const isMergeAnchor = !!merge && merge.top === row && merge.left === col

  const floatingImages: Array<{ index: number; anchor: ImageAnchor }> = []
  for (let i = 0; i < sheet.images.length; i++) {
    const a = sheet.images[i]
    if (imageAnchorContains(a, row, col)) floatingImages.push({ index: i, anchor: { ...a } })
  }

  const dispImgId = live?.dispImgId
  const ci = dispImgId ? workbook.cellImages?.get(dispImgId) : undefined
  const cellImage = ci ? { id: dispImgId!, src: ci.src, mime: ci.mime } : null

  const dataValidation = sheet.dataValidations.find(
    (r) => row >= r.top && row <= r.bottom && col >= r.left && col <= r.right,
  ) ?? null

  const engine = new ConditionalEngine(sheet)
  const hits = engine.hasRules() ? engine.inspectHits(row, col, snapshot.raw) : []
  const conditional = hits.map((h) => ({
    ruleIndex: h.ruleIndex,
    style: effectToStyleOverride(h.effect),
  }))

  return {
    ...snapshot,
    merge: merge ? { ...merge } : null,
    isMergeAnchor,
    floatingImages,
    cellImage,
    dataValidation: dataValidation ? { ...dataValidation } : null,
    conditional,
    hyperlink: live?.hyperlink ?? null,
    comment: live?.comment ?? null,
  }
}

/** CellEffect(渲染层用)→ CellStyleOverride(对外 API 用) */
function effectToStyleOverride(effect: import('../render/conditional').CellEffect): CellStyleOverride {
  const out: CellStyleOverride = {}
  if (effect.fillColor) out.fill = { type: 'solid', fgColor: effect.fillColor }
  if (effect.fontColor || effect.bold) out.font = { ...(effect.fontColor ? { color: effect.fontColor } : {}), ...(effect.bold ? { bold: true } : {}) }
  return out
}
