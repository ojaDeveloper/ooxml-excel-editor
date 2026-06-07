/**
 * 模板填值(P3)—— 把一份真实 .xlsx 当模板,JSON 数据按两种方式注入:
 *
 *  1. **占位符 `{{key}}`** —— 扫全表 `cell.type==='string'` 的格,文本里出现 {{name}} / {{a.b.c}}
 *     被 `placeholders` 字典里的值替换。未匹配的占位符保留原样,不报错。dot path 支持嵌套对象。
 *
 *  2. **锚点表(`startCell + rows`)** —— 从指定单元格起按位铺二维数组 / 对象数组;
 *     对象数组按 `columns` 顺序(没给用首行 keys 顺序)。
 *
 * **不入命令栈**(预渲染处理,不算编辑;ViewerController.applyTemplate 触发 onModelChange 重渲)。
 * 进度 + AbortSignal:每 100 格 emit + yield。
 */
import type { CellModel, SheetModel, WorkbookModel } from '../model/types'
import { cellKey } from '../model/types'
import { setCellValue } from '../model/mutations'
import type { ExportProgressFn } from '../progress'
import { checkAborted, yieldToEvent } from '../export/abort'

export interface TemplateAnchor {
  /** 起始单元格:'A5' 或 {row, col} */
  startCell: string | { row: number; col: number }
  /** 数据行:二维数组(按位置铺)或对象数组(按 columns 顺序) */
  rows: unknown[][] | Record<string, unknown>[]
  /** 对象数组列名顺序(不给用首行 keys 顺序) */
  columns?: string[]
  /** 落在哪张表(缺省 = activeSheet) */
  sheetName?: string
  /**
   * **未填行清理**(默认 `true`)。模板里常预留 N 行带边框 / 样式的"空白数据区"等 JSON 填进来,
   * 当 JSON 行数 < N 时,多余的空白行会显示成"看起来像数据"的边框。开启后:从最后一行已填行往下扫,
   * 直到撞到首个**含 raw 值**的格(如 `{{total}}` 替换后的总计行)为止,把这段空白行中**锚点列范围**
   * 内的格全清掉(`sheet.cells.delete`)。**只清锚点列、不动模板的其他列与超出范围的行**。
   *
   * 关掉(`false`)= 保留模板原样(连边框带空格一起渲染)。
   */
  trimUnused?: boolean
}

export interface TemplateFillSpec {
  /** 占位符字典,支持 {{path.to.value}} dot path */
  placeholders?: Record<string, unknown>
  /** 锚点表(可多个,可跨表) */
  anchors?: TemplateAnchor[]
  /** 长任务进度回调 */
  onProgress?: ExportProgressFn
  /** 取消信号 */
  signal?: AbortSignal
}

/** "A1" / "AB12" → {row, col}(0-based);bad → null */
export function parseCellAddress(addr: string): { row: number; col: number } | null {
  const m = /^([A-Z]+)(\d+)$/.exec(addr.trim().toUpperCase())
  if (!m) return null
  let col = 0
  for (let i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64)
  return { row: parseInt(m[2], 10) - 1, col: col - 1 }
}

/** 解析 {{a.b.c}} dot path */
function getByPath(obj: Record<string, unknown>, path: string): unknown {
  const segs = path.split('.')
  let cur: unknown = obj
  for (const s of segs) {
    if (cur == null || typeof cur !== 'object') return undefined
    cur = (cur as Record<string, unknown>)[s]
  }
  return cur
}

/** 把字符串里 {{key}} / {{a.b}} 替换为 placeholders 的值;缺失 token 保留原样 */
export function replacePlaceholders(text: string, placeholders: Record<string, unknown>): string {
  return text.replace(/\{\{\s*([\w.]+)\s*\}\}/g, (full, path: string) => {
    const v = getByPath(placeholders, path)
    return v == null ? full : String(v)
  })
}

/** 在 sheet 上扫所有 string 格,替换文本里的 {{key}}。返回处理过的格数。 */
async function replacePlaceholdersInSheet(
  sheet: SheetModel,
  placeholders: Record<string, unknown>,
  signal?: AbortSignal,
  onProgress?: ExportProgressFn,
  sheetIndex = 0,
): Promise<number> {
  let done = 0
  for (const cell of sheet.cells.values()) {
    if (cell.type !== 'string' || typeof cell.raw !== 'string') continue
    const next = replacePlaceholders(cell.raw, placeholders)
    if (next !== cell.raw) cell.raw = next
    done++
    if ((done % 100) === 0) {
      checkAborted(signal)
      onProgress?.({ stage: 'render', sheetIndex, ratio: undefined, label: `占位符替换…(已扫 ${done})` })
      await yieldToEvent()
    }
  }
  return done
}

/** 在 sheet 上从 startCell 起铺 anchor.rows;复用 mutations.setCellValue 推断类型 + 维度自增。
 *  默认 `trimUnused: true` 清理"模板预留但 JSON 未填"的空白行(防边框看起来像数据)。 */
function applyAnchor(sheet: SheetModel, anchor: TemplateAnchor): number {
  const at = typeof anchor.startCell === 'string' ? parseCellAddress(anchor.startCell) : anchor.startCell
  if (!at) return 0
  const rows = anchor.rows
  if (!rows.length) return 0
  const first = rows[0]
  const isObjectArray = first !== null && typeof first === 'object' && !Array.isArray(first)
  const keys: string[] = isObjectArray ? (anchor.columns ?? Object.keys(first as Record<string, unknown>)) : []
  let written = 0
  let maxColOffset = 0 // 记录最大列偏移,trim 时只清理"锚点用到的列"
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]
    const arr: unknown[] = Array.isArray(row) ? row : keys.map((k) => (row as Record<string, unknown>)[k])
    if (arr.length - 1 > maxColOffset) maxColOffset = arr.length - 1
    for (let c = 0; c < arr.length; c++) {
      const r = at.row + i
      const col = at.col + c
      const v = arr[c]
      if (v == null || v === '') continue
      // setCellValue 走"对外语义"路径,类型自动推断 + 维度自增 + dispImgId 清理等
      setCellValue(sheet, r, col, v as string | number | boolean | Date)
      written++
    }
  }

  // P3 进阶:trim 未填行(默认 on)。从首个未填行往下扫,遇到第一个**含 raw 值**的格停;
  // 这段空白行内锚点列范围(at.col..at.col+maxColOffset)里的格全清掉。
  if (anchor.trimUnused !== false) {
    const startCol = at.col
    const endCol = at.col + maxColOffset
    const firstUnusedRow = at.row + rows.length
    for (let r = firstUnusedRow; r < sheet.dimension.rows; r++) {
      let hasContent = false
      for (let c = startCol; c <= endCol; c++) {
        const cell = sheet.cells.get(cellKey(r, c))
        if (cell && cell.raw != null && cell.raw !== '') {
          hasContent = true
          break
        }
      }
      if (hasContent) break // 撞到"真"内容(如 {{total}} 已替换的合计行)→ 停
      // 清这一段空白行内锚点列的格(只是边框/样式占位)
      for (let c = startCol; c <= endCol; c++) {
        sheet.cells.delete(cellKey(r, c))
      }
    }
  }
  return written
}

/**
 * 主入口:**原地**修改 workbook(注意:壳调用方需自己 clone 才能保留模板原态)。
 * 返回处理结果概要(各阶段计数)。
 */
export async function fillTemplate(
  workbook: WorkbookModel,
  spec: TemplateFillSpec,
): Promise<{ placeholdersScanned: number; anchorsWritten: number }> {
  checkAborted(spec.signal)
  // 1. 占位符替换
  let placeholdersScanned = 0
  if (spec.placeholders) {
    for (let i = 0; i < workbook.sheets.length; i++) {
      checkAborted(spec.signal)
      placeholdersScanned += await replacePlaceholdersInSheet(
        workbook.sheets[i], spec.placeholders, spec.signal, spec.onProgress, i,
      )
    }
  }

  // 2. 锚点表
  let anchorsWritten = 0
  if (spec.anchors) {
    for (let i = 0; i < spec.anchors.length; i++) {
      checkAborted(spec.signal)
      const a = spec.anchors[i]
      const sheet = a.sheetName ? workbook.sheets.find((s) => s.name === a.sheetName) : workbook.sheets[workbook.activeSheet]
      if (!sheet) continue
      anchorsWritten += applyAnchor(sheet, a)
      spec.onProgress?.({ stage: 'render', sheetIndex: sheet.index, ratio: (i + 1) / spec.anchors.length, label: `锚点 ${i + 1}/${spec.anchors.length}` })
      await yieldToEvent()
    }
  }
  return { placeholdersScanned, anchorsWritten }
}

/** 为防 cell.raw 类型 issue,暴露给外部的小工具:取 cellKey(测试用) */
export { cellKey }
export type { CellModel }
