/**
 * 复制 1:1 保真快照(同应用内 + 跨实例 Vue3/Vue2/React 互相复制都一致)。
 *
 * 普通剪贴板走 text/html(给 Excel/WPS/Word)+ text/plain,但 HTML 是有损交换:
 * cellInlineStyle 只带 粗斜/字色/底色/对齐,丢边框、数字格式、原始值(数字会变成格式化后的文本)、
 * 行高列宽、合并、图片。要 1:1,必须把**完整模型快照**序列化进剪贴板本身 —— 嵌在 HTML 的
 * `data-ooxml-clip` 属性里(base64(UTF-8 JSON))。任何本组件实例读 text/html 都能反序列化 1:1 还原;
 * 外部应用(Excel)忽略这个属性,只读可见 `<table>`。跨标签页/跨实例因此也能 1:1(快照随剪贴板走,
 * 不依赖内存)。
 *
 * 携带:每格(原始值/类型/公式/超链/批注/富文本/dispImgId + 完整 CellStyle)、合并、浮动图(base64)、
 * DISPIMG 单元格图字节(base64,落地时登记进目标 cellImages)、行高、列宽。坐标全部相对复制区左上。
 */
import type { AnchorCell, CellModel, CellStyle, CellValueType, MergeRange, RichTextRun, SheetModel, WorkbookModel } from '../model/types'
import { cellKey, makeDefaultStyle } from '../model/types'

/** raw 是 Date 时序列化为 {__d: epochMs},反序列化还原。其余原样 JSON。 */
type ClipRaw = number | string | boolean | null | { __d: number }

export interface ClipCell {
  r: number // 相对复制区左上的行偏移
  c: number // 列偏移
  type: CellValueType
  raw: ClipRaw
  formula?: string
  hyperlink?: string
  comment?: string
  rich?: RichTextRun[]
  dispImgId?: string
  style: CellStyle
}
export interface ClipImage {
  from: AnchorCell // 相对锚点
  to?: AnchorCell
  extWidthEmu?: number
  extHeightEmu?: number
  editAs?: string
  mime: string
  b64: string
}
export interface ClipCellImage { id: string; mime: string; b64: string }
export interface ClipDim { i: number; height?: number; width?: number; hidden?: boolean; custom?: boolean }

export interface ClipSnapshot {
  v: 1
  rows: number
  cols: number
  cells: ClipCell[]
  merges: MergeRange[] // 相对
  images: ClipImage[]
  cellImages: ClipCellImage[]
  rowHeights: ClipDim[] // i = 相对行
  colWidths: ClipDim[] // i = 相对列
}

/** 图片字节预算(原始字节):复制区图片总字节超此值 → 降级为"无图 1:1 复制",避免剪贴板超限/卡顿。 */
export const CLIP_IMAGE_BUDGET_BYTES = 6 * 1024 * 1024


/**
 * 抓一段区域的完整模型快照(相对坐标)。
 * `withImageBytes=false`:图片只记引用(id / 浮动序号),b64 留空 —— 给"图片字节走可见 `<img>`、
 * 快照只引用"的瘦身传输用(避免图片被双重 base64),粘贴时由 parseSnapshotHtml 从 `<img>` 回填字节。
 */
export function serializeSnapshot(sheet: SheetModel, wb: WorkbookModel, range: MergeRange, opts: { withImageBytes?: boolean } = {}): ClipSnapshot {
  const withBytes = opts.withImageBytes !== false
  const rows = range.bottom - range.top + 1
  const cols = range.right - range.left + 1
  const cells: ClipCell[] = []
  const usedCellImageIds = new Set<string>()
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cell = sheet.cells.get(cellKey(range.top + r, range.left + c))
      if (!cell) continue
      const style = sheet.styles[cell.styleId] ?? sheet.styles[0] ?? makeDefaultStyle()
      const cc: ClipCell = {
        r,
        c,
        type: cell.type,
        raw: cell.raw instanceof Date ? { __d: cell.raw.getTime() } : (cell.raw as ClipRaw),
        style,
      }
      if (cell.formula) cc.formula = cell.formula
      if (cell.hyperlink) cc.hyperlink = cell.hyperlink
      if (cell.comment) cc.comment = cell.comment
      if (cell.rich) cc.rich = cell.rich
      if (cell.dispImgId) { cc.dispImgId = cell.dispImgId; usedCellImageIds.add(cell.dispImgId) }
      cells.push(cc)
    }
  }
  const merges: MergeRange[] = []
  for (const m of sheet.merges) {
    if (m.top >= range.top && m.left >= range.left && m.bottom <= range.bottom && m.right <= range.right) {
      merges.push({ top: m.top - range.top, left: m.left - range.left, bottom: m.bottom - range.top, right: m.right - range.left })
    }
  }
  const images: ClipImage[] = []
  for (const im of sheet.images) {
    if (im.from.row < range.top || im.from.row > range.bottom || im.from.col < range.left || im.from.col > range.right) continue
    if (!im.bytes || !im.mime) continue // blob-only 无字节的图无法序列化(跟导出一致,跳过)
    images.push({
      from: { col: im.from.col - range.left, row: im.from.row - range.top, colOffEmu: im.from.colOffEmu, rowOffEmu: im.from.rowOffEmu },
      to: im.to ? { col: im.to.col - range.left, row: im.to.row - range.top, colOffEmu: im.to.colOffEmu, rowOffEmu: im.to.rowOffEmu } : undefined,
      extWidthEmu: im.extWidthEmu,
      extHeightEmu: im.extHeightEmu,
      editAs: im.editAs,
      mime: im.mime,
      b64: withBytes ? bytesToB64(im.bytes) : '',
    })
  }
  const cellImages: ClipCellImage[] = []
  for (const id of usedCellImageIds) {
    const ci = wb.cellImages?.get(id)
    if (ci?.bytes && ci.mime) cellImages.push({ id, mime: ci.mime, b64: withBytes ? bytesToB64(ci.bytes) : '' })
  }
  const rowHeights: ClipDim[] = []
  for (let r = 0; r < rows; r++) {
    const info = sheet.rows.get(range.top + r)
    if (info) rowHeights.push({ i: r, height: info.height, hidden: info.hidden || undefined, custom: info.customHeight || undefined })
  }
  const colWidths: ClipDim[] = []
  for (let c = 0; c < cols; c++) {
    const info = sheet.columns.get(range.left + c)
    if (info) colWidths.push({ i: c, width: info.width, hidden: info.hidden || undefined })
  }
  return { v: 1, rows, cols, cells, merges, images, cellImages, rowHeights, colWidths }
}

/** ClipRaw → 运行时 raw(还原 Date)。 */
export function reviveClipRaw(raw: ClipRaw): CellModel['raw'] {
  if (raw && typeof raw === 'object' && '__d' in raw) return new Date(raw.__d)
  return raw
}

export function encodeSnapshot(snap: ClipSnapshot): string {
  return utf8ToB64(JSON.stringify(snap))
}

export function decodeSnapshot(s: string | null | undefined): ClipSnapshot | null {
  if (!s) return null
  try {
    const o = JSON.parse(b64ToUtf8(s)) as ClipSnapshot
    return o && o.v === 1 && Array.isArray(o.cells) ? o : null
  } catch {
    return null
  }
}

/**
 * 把图片字节从可见 `<img data-clip-img="key">` 回填进瘦身快照(传输优化:图片只在 `<img>` 存一份)。
 * key 约定:DISPIMG 单元格图 = `c:${id}`,浮动图 = `f:${序号}`(序号同 serializeSnapshot 的 images 顺序)。
 * 已带 b64(完整快照路径)则原样保留。
 */
export function reattachImages(snap: ClipSnapshot | null, imgB64: Map<string, string>): ClipSnapshot | null {
  if (!snap) return null
  return {
    ...snap,
    cellImages: snap.cellImages.map((ci) => ({ ...ci, b64: ci.b64 || imgB64.get(`c:${ci.id}`) || '' })),
    images: snap.images.map((im, i) => ({ ...im, b64: im.b64 || imgB64.get(`f:${i}`) || '' })),
  }
}

/** 降级:去掉所有图片(超字节预算时用)。DISPIMG 格中性化为空格(保样式),避免粘出引用不到的破图。 */
export function withoutImages(snap: ClipSnapshot): ClipSnapshot {
  return {
    ...snap,
    cells: snap.cells.map((c) => (c.dispImgId ? { ...c, type: 'empty', raw: null, formula: undefined, dispImgId: undefined } : c)),
    images: [],
    cellImages: [],
  }
}

/** 从剪贴板 HTML 抽出本组件的 1:1 快照(`<table data-ooxml-clip="...">` + 回填 `<img>` 字节);非本组件复制返 null。 */
export function parseSnapshotHtml(html: string): ClipSnapshot | null {
  if (typeof DOMParser === 'undefined' || !html) return null
  try {
    const doc = new DOMParser().parseFromString(html, 'text/html')
    const snap = decodeSnapshot(doc.querySelector('table')?.getAttribute('data-ooxml-clip'))
    if (!snap) return null
    const imgB64 = new Map<string, string>()
    doc.querySelectorAll('img[data-clip-img]').forEach((img) => {
      const key = img.getAttribute('data-clip-img')
      const src = img.getAttribute('src') || ''
      const comma = src.indexOf(',')
      if (key && src.startsWith('data:') && comma > 0) imgB64.set(key, src.slice(comma + 1))
    })
    return reattachImages(snap, imgB64)
  } catch {
    return null
  }
}

// ---- base64 helpers(UTF-8 安全 + 二进制) ----
export function bytesToB64(bytes: Uint8Array): string {
  let bin = ''
  const CHUNK = 0x8000
  for (let i = 0; i < bytes.length; i += CHUNK) bin += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
  return btoa(bin)
}
export function b64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}
function utf8ToB64(s: string): string {
  return bytesToB64(new TextEncoder().encode(s))
}
function b64ToUtf8(b64: string): string {
  return new TextDecoder().decode(b64ToBytes(b64))
}
