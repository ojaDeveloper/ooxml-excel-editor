/**
 * 模型变更(纯函数,框架无关)—— data-access 的"写"另一半。
 * 只改 SheetModel,不碰渲染;调用方负责重绘(同 sortColumn 的"改完重绘")。
 * 全部以"前态可逆"为原则:返回旧值供命令栈做 undo。
 */
import type { CellImage, CellModel, CellStyle, CellStyleOverride, ColumnInfo, ImageAnchor, MergeRange, RowInfo, SheetModel, WorkbookModel } from './types'
import { cellKey } from './types'
import type { CellValue } from './data-access'
import { pxToEmu } from '../layout/units'

/** 由输入值推断单元格类型 + 组装 CellModel(轻量 Excel 式输入:= 开头 → 公式,纯数字串 → 数字)。 */
function inferCell(row: number, col: number, value: CellValue, styleId: number): CellModel | null {
  if (value === null || value === undefined || value === '') return null // 空 → 删除
  if (typeof value === 'number') return { row, col, type: 'number', raw: value, styleId }
  if (typeof value === 'boolean') return { row, col, type: 'boolean', raw: value, styleId }
  if (value instanceof Date) return { row, col, type: 'date', raw: value, styleId }
  // string:
  const s = value
  if (s[0] === '=') return { row, col, type: 'formula', raw: null, formula: s, styleId } // E4 前无重算,raw 待引擎填
  if (s.trim() !== '' && !isNaN(Number(s))) return { row, col, type: 'number', raw: Number(s), styleId }
  return { row, col, type: 'string', raw: s, styleId }
}

function growDimension(sheet: SheetModel, row: number, col: number): void {
  if (row + 1 > sheet.dimension.rows) sheet.dimension.rows = row + 1
  if (col + 1 > sheet.dimension.cols) sheet.dimension.cols = col + 1
}

/** 设单元格值。返回旧 CellModel(克隆,供 undo);新值为空则删除该格。 */
export function setCellValue(sheet: SheetModel, row: number, col: number, value: CellValue): void {
  const key = cellKey(row, col)
  const prev = sheet.cells.get(key)
  const styleId = prev?.styleId ?? 0
  const next = inferCell(row, col, value, styleId)
  if (next) {
    sheet.cells.set(key, next)
    growDimension(sheet, row, col)
  } else {
    sheet.cells.delete(key)
  }
}

/** 清空单元格(删除该格,保留 dimension)。 */
export function clearCell(sheet: SheetModel, row: number, col: number): void {
  sheet.cells.delete(cellKey(row, col))
}

/** 设/清单元格批注(1.11.0):空批注清除;格不存在且有批注 → 建空格挂批注。 */
export function setCellComment(sheet: SheetModel, row: number, col: number, comment: string): void {
  const key = cellKey(row, col)
  const cell = sheet.cells.get(key)
  const text = comment.trim()
  if (cell) {
    if (text) cell.comment = text
    else {
      delete cell.comment
      // 空格(无值/无批注/无可见样式)清掉,避免膨胀;有值/样式的格保留
      if (cell.type === 'empty' && !cell.comment) sheet.cells.delete(key)
    }
  } else if (text) {
    sheet.cells.set(key, { row, col, type: 'empty', raw: null, comment: text, styleId: 0 })
    growDimension(sheet, row, col)
  }
}

/** 区域批量设值(2D,左上对齐 range.top/left)。 */
export function setRangeValues(sheet: SheetModel, range: MergeRange, values: CellValue[][]): void {
  for (let r = 0; r < values.length; r++) {
    const rowArr = values[r]
    for (let c = 0; c < rowArr.length; c++) {
      setCellValue(sheet, range.top + r, range.left + c, rowArr[c])
    }
  }
}

/**
 * 写回公式引擎算出的结果(只改 raw,保留 formula/type/styleId/rich)。
 * 公式格:raw = 计算值(显示用 raw);被公式引用的值格:不变(raw 已是其值)。空格跳过。
 */
export function setCellComputed(sheet: SheetModel, row: number, col: number, value: CellValue): void {
  const cell = sheet.cells.get(cellKey(row, col))
  if (!cell) return
  cell.raw = value
}

/** 直接写回/删除一个格的底层 CellModel(命令逆向用:精确还原前态)。 */
export function restoreCell(sheet: SheetModel, row: number, col: number, cell: CellModel | null): void {
  const key = cellKey(row, col)
  if (cell) {
    sheet.cells.set(key, cell)
    growDimension(sheet, row, col)
  } else {
    sheet.cells.delete(key)
  }
}

// ====================== 维度(列宽/行高)写 ======================
// 注:存储单位与 renderer 一致——列宽=px(已从字符换算)、行高=px(已从 pt 换算)。
// 这里只写模型;调用方负责 rebuildMetrics + 重绘(同 setCellValue 的"改完重绘")。

/** 设列宽(px,非缩放存储);保留 hidden。 */
export function setColumnWidth(sheet: SheetModel, col: number, width: number): void {
  const info = sheet.columns.get(col)
  sheet.columns.set(col, { width: Math.max(8, width), hidden: info?.hidden ?? false })
}

/** 设行高(px,非缩放存储);保留 hidden。 */
export function setRowHeight(sheet: SheetModel, row: number, height: number): void {
  const info = sheet.rows.get(row)
  sheet.rows.set(row, { height: Math.max(6, height), hidden: info?.hidden ?? false })
}

// ====================== 图片(浮动/嵌入)写 ======================
// 作用 SheetModel.images: ImageAnchor[]。移动/缩放统一规整成"原点相对"oneCellAnchor
// (from.col/row=0 + colOff/rowOff EMU = 距原点偏移,ext = 尺寸),避免跨列/行进位的换算。

/** 浅克隆 ImageAnchor(共享不可变的 bytes/src,省内存;命令逆向 / 拖拽起始态用)。 */
export function cloneImageAnchor(a: ImageAnchor): ImageAnchor {
  return { ...a, from: { ...a.from }, to: a.to ? { ...a.to } : undefined }
}

/** 加一张图(index 缺省追加到末尾),返回插入位置。 */
export function addImage(sheet: SheetModel, anchor: ImageAnchor, index?: number): number {
  const at = index ?? sheet.images.length
  sheet.images.splice(at, 0, anchor)
  return at
}

/** 删一张图(调用方负责为 undo 捕获前态)。 */
export function removeImage(sheet: SheetModel, index: number): void {
  sheet.images.splice(index, 1)
}

/** 整个 CellModel 落格(承载公式/超链/批注/富文本/dispImgId 等 setCellValue 推不出的字段)+ 撑维度。 */
export function setCellModel(sheet: SheetModel, cell: CellModel): void {
  sheet.cells.set(cellKey(cell.row, cell.col), cell)
  growDimension(sheet, cell.row, cell.col)
}

// ---- WPS 单元格内嵌图(DISPIMG)⇄ 浮动图 互转(第二期) ----

/** 为登记表生成一个未占用的 DISPIMG id(确定性:从计数找空位,不依赖随机) */
function freshCellImageId(reg: Map<string, CellImage>): string {
  let n = reg.size
  let id = `ID_cell_${n}`
  while (reg.has(id)) id = `ID_cell_${++n}`
  return id
}

/** 全簿是否还有单元格引用某 DISPIMG id(用于决定登记项可否回收) */
function anyCellRefs(wb: WorkbookModel, id: string): boolean {
  for (const sheet of wb.sheets) for (const c of sheet.cells.values()) if (c.dispImgId === id) return true
  return false
}

/**
 * 浮动图 → WPS 单元格内嵌图(DISPIMG)。
 * 取 sheet.images[imageIndex] 的字节登记到 wb.cellImages(新 id),目标格设 =DISPIMG 公式 + dispImgId,
 * 移除该浮动图。返回新 id;图缺 bytes/mime 不可转 → 返 null(不动模型)。
 */
export function convertFloatToCellImage(
  wb: WorkbookModel,
  sheet: SheetModel,
  imageIndex: number,
  row: number,
  col: number,
): string | null {
  const img = sheet.images[imageIndex]
  if (!img || !img.bytes || !img.mime) return null
  if (!wb.cellImages) wb.cellImages = new Map()
  const id = freshCellImageId(wb.cellImages)
  wb.cellImages.set(id, { id, bytes: img.bytes, mime: img.mime, src: img.src || '' })
  const existing = sheet.cells.get(cellKey(row, col))
  sheet.cells.set(cellKey(row, col), {
    row,
    col,
    type: 'formula',
    raw: null,
    formula: `_xlfn.DISPIMG("${id}",1)`,
    dispImgId: id,
    styleId: existing?.styleId ?? 0,
  })
  growDimension(sheet, row, col)
  sheet.images.splice(imageIndex, 1)
  // 原本是空格(渲染为白)→ 别让默认 styleId 0 的填充冒出来(有些文件 styles[0] 自带灰底如 #D0CECE),
  // 清掉填充保持原来的白。原本有样式的格保留不动(绝不主动改色)。
  if (!existing) applyStyleOverride(sheet, row, col, { fill: { type: 'none' } })
  return id
}

/**
 * WPS 单元格内嵌图 → 浮动图。
 * 取 (row,col) 的 dispImgId 对应登记图,造一张 oneCellAnchor 浮动图锚在该格(默认尺寸 sizePx),
 * 清空该格。登记表里若无其它格再引用该 id,一并回收。返回浮动图索引;非内嵌图格返 -1。
 */
export function convertCellImageToFloat(
  wb: WorkbookModel,
  sheet: SheetModel,
  row: number,
  col: number,
  size: { width: number; height: number } = { width: 96, height: 96 },
): number {
  const key = cellKey(row, col)
  const cell = sheet.cells.get(key)
  const id = cell?.dispImgId
  if (!id) return -1
  const ci = wb.cellImages?.get(id)
  const anchor: ImageAnchor = {
    src: ci?.src ?? '',
    bytes: ci?.bytes,
    mime: ci?.mime,
    from: { col, row, colOffEmu: 0, rowOffEmu: 0 },
    extWidthEmu: pxToEmu(size.width),
    extHeightEmu: pxToEmu(size.height),
    editAs: 'oneCell',
  }
  sheet.cells.delete(key) // 内嵌图本占满格,转浮动后该格清空
  const at = sheet.images.length
  sheet.images.push(anchor)
  if (!anyCellRefs(wb, id)) wb.cellImages?.delete(id)
  return at
}

/**
 * 设图片的内容矩形(zoom 后像素,content 坐标)→ 规整为原点相对 oneCellAnchor。
 * 像素↔EMU:left_px = emuToPx(colOffEmu)*zoom ⇒ colOffEmu = pxToEmu(left/zoom)。不依赖列/行几何。
 */
export function setImageRect(
  sheet: SheetModel,
  index: number,
  rect: { left: number; top: number; width: number; height: number },
  zoom: number,
): void {
  const img = sheet.images[index]
  if (!img) return
  const toEmu = (px: number) => pxToEmu(px / zoom)
  img.from = { col: 0, row: 0, colOffEmu: toEmu(rect.left), rowOffEmu: toEmu(rect.top) }
  img.to = undefined
  img.extWidthEmu = toEmu(rect.width)
  img.extHeightEmu = toEmu(rect.height)
}

/** 直接写回/删除一个列/行的维度信息(命令逆向用:精确还原前态;null=删 Map 项回落默认宽高)。 */
export function restoreDimension(
  sheet: SheetModel,
  axis: 'col' | 'row',
  index: number,
  info: ColumnInfo | RowInfo | null,
): void {
  const map = axis === 'col' ? sheet.columns : sheet.rows
  if (info) map.set(index, info as ColumnInfo & RowInfo)
  else map.delete(index)
}

/** 样式去重入表,返回 styleId(E5 样式编辑用;深比较已存在则复用)。 */
export function internStyle(sheet: SheetModel, style: CellStyle): number {
  const json = JSON.stringify(style)
  for (let i = 0; i < sheet.styles.length; i++) if (JSON.stringify(sheet.styles[i]) === json) return i
  sheet.styles.push(style)
  return sheet.styles.length - 1
}

/** 合并样式覆盖到完整样式: font/fill/borders 浅合并,其余字段覆盖(单一真相源,渲染与编辑共用)。 */
export function mergeStyleOverride(base: CellStyle, over: CellStyleOverride): CellStyle {
  return {
    ...base,
    ...over,
    font: over.font ? { ...base.font, ...over.font } : base.font,
    fill: over.fill ? { ...base.fill, ...over.fill } : base.fill,
    borders: over.borders ? { ...base.borders, ...over.borders } : base.borders,
  }
}

/**
 * 给一个格套样式覆盖(E5):取基样式 → 合并 → intern → 改 styleId。
 * 空格(不在 Map)→ 以默认样式为基,新建一个 type='empty' 的格承载 styleId(像 Excel 给空格上色)。
 */
export function applyStyleOverride(sheet: SheetModel, row: number, col: number, patch: CellStyleOverride): void {
  const key = cellKey(row, col)
  const cell = sheet.cells.get(key)
  const base = (cell ? sheet.styles[cell.styleId] : sheet.styles[0]) ?? sheet.styles[0]
  if (!base) return // 无任何样式表(极端防御)
  const styleId = internStyle(sheet, mergeStyleOverride(base, patch))
  if (cell) {
    cell.styleId = styleId
  } else {
    sheet.cells.set(key, { row, col, type: 'empty', raw: null, styleId })
    growDimension(sheet, row, col)
  }
}
