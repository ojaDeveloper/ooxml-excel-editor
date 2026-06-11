/**
 * 粘贴行为配置(框架无关)。控制 Ctrl+V / 右键粘贴时,源内容的各方面如何落到目标:
 * 覆盖目标原有(贴近源)、合并(源没写的保留目标)、还是只填值不动样式/结构。
 *
 * 默认 = 用户约定的"覆盖式 1:1"(见各字段默认值);不配置即走默认。
 * 右键「选择性粘贴 ▸ 保留原来的(仅值)」走 {@link PASTE_PRESET_VALUES_ONLY} 预设。
 * 可经 `viewer.setPasteBehavior(cfg)` / 组件 `:paste-behavior` prop 改默认。
 *
 * 首行 vs 中间唯一的差异收敛在 `colWidth: 'firstRowOnly'`(列宽整列共享,仅粘到首行 start.row===0 才搬源,
 * 避免粘到中间把上方表头列宽改掉);其余各项首行/中间一视同仁。
 */
import type { CellStyle, CellStyleOverride, Fill } from '../model/types'

export interface PasteBehavior {
  /** 字体/对齐/换行/边框/数字格式: 'overwrite' 覆盖式(只留源) | 'merge' 合并式(源没写的留目标) | 'skip' 不动 */
  cellStyle: 'overwrite' | 'merge' | 'skip'
  /** 填充底色: 'overwrite' 覆盖式(源没写→无填充) | 'merge' 合并式(源没写→留目标底色) | 'skip' 不动 */
  fill: 'overwrite' | 'merge' | 'skip'
  /** 行高: 'source' 搬源行高 | 'keep' 不动 */
  rowHeight: 'source' | 'keep'
  /** 列宽: 'source' 总搬源 | 'keep' 不动 | 'firstRowOnly' 仅粘到首行(start.row===0)搬源 */
  colWidth: 'source' | 'keep' | 'firstRowOnly'
  /** 源自带的合并区: 'apply' 应用 | 'skip' 不应用 */
  sourceMerges: 'apply' | 'skip'
  /** 目标原有、落在粘贴区内的合并: 'clear' 清掉(否则旧合并会吞列致数据错位) | 'keep' 保留 */
  targetMerges: 'clear' | 'keep'
  /** 图片(内嵌/浮动): 'apply' 落格 | 'skip' 不粘 */
  images: 'apply' | 'skip'
}

/** 默认粘贴行为 = 覆盖式 1:1(贴近源)。不配置即此。 */
export const DEFAULT_PASTE_BEHAVIOR: PasteBehavior = {
  cellStyle: 'overwrite',
  fill: 'overwrite',
  rowHeight: 'source',
  colWidth: 'firstRowOnly',
  sourceMerges: 'apply',
  targetMerges: 'clear',
  images: 'apply',
}

/** 右键预设:保留原来的(仅值)—— 只落值,样式/底色/行高列宽/合并/图片全不动,目标结构完全保留。 */
export const PASTE_PRESET_VALUES_ONLY: PasteBehavior = {
  cellStyle: 'skip',
  fill: 'skip',
  rowHeight: 'keep',
  colWidth: 'keep',
  sourceMerges: 'skip',
  targetMerges: 'keep',
  images: 'skip',
}

/** 补全部分配置为完整 PasteBehavior(缺项回落默认)。 */
export function resolvePasteBehavior(partial?: Partial<PasteBehavior> | null): PasteBehavior {
  return { ...DEFAULT_PASTE_BEHAVIOR, ...(partial ?? {}) }
}

/** 列宽该不该搬:source 总搬;firstRowOnly 仅 start.row===0 搬;keep 不搬。 */
export function shouldApplyColWidth(b: PasteBehavior, startRow: number): boolean {
  return b.colWidth === 'source' || (b.colWidth === 'firstRowOnly' && startRow === 0)
}

/**
 * 按 cellStyle / fill 两档模式,从「目标现有样式 target + 源 patch」算出粘贴后该格的完整 CellStyle。
 * 返回 null = 两档都 skip,样式整个不动(调用方跳过)。neutral = 表的中性默认(styles[0])。
 *
 * 非填充部分(font/对齐/换行/边框/数字格式):overwrite 以 neutral 为基套源(只留源);merge 以 target 为基套源
 *   (源没写的留目标);skip 保留 target。填充部分(fill):overwrite 源没写→无填充;merge 源没写→留目标;skip 留目标。
 */
export function resolvePastedCellStyle(
  target: CellStyle,
  neutral: CellStyle,
  patch: CellStyleOverride,
  styleMode: PasteBehavior['cellStyle'],
  fillMode: PasteBehavior['fill'],
): CellStyle | null {
  if (styleMode === 'skip' && fillMode === 'skip') return null
  const { fill: patchFill, ...patchNonFill } = patch
  const base = styleMode === 'overwrite' ? neutral : target
  const result: CellStyle = { ...base }
  if (styleMode !== 'skip') {
    Object.assign(result, patchNonFill) // hAlign/vAlign/wrapText/numFmt/shrinkToFit/textRotation/indent 覆盖
    if (patchNonFill.font) result.font = { ...base.font, ...patchNonFill.font }
    if (patchNonFill.borders) result.borders = { ...base.borders, ...patchNonFill.borders }
  } else {
    result.font = { ...base.font }
    result.borders = { ...base.borders }
  }
  // 填充
  if (fillMode === 'skip') result.fill = target.fill
  else if (patchFill) result.fill = { ...(patchFill as Fill) }
  else result.fill = fillMode === 'overwrite' ? { type: 'none' } : target.fill
  return result
}
