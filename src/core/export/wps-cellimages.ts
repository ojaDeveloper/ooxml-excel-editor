/**
 * WPS 单元格内嵌图(DISPIMG)导出回注 —— 第三期:解决"导出回 .xlsx 丢内嵌图"。
 *
 * ExcelJS 不建模 `xl/cellimages.xml` 这个 WPS 私有件,load→write 会把它(及 workbook 对它的引用、
 * Content_Types 声明)整个丢掉。本模块在 ExcelJS 写出 zip **之后**做后处理,从当前 WorkbookModel.cellImages
 * 重新生成并回注全套零件,使导出的 .xlsx 在 WPS 里仍能显示单元格内嵌图(含 App 内"浮动→嵌入"新转的图)。
 *
 * 回注的零件(齐全才能被 WPS 识别):
 *   - xl/media/cellimageN.{ext}        图片字节
 *   - xl/cellimages.xml                登记表(cNvPr@name = DISPIMG id,blip@embed → rels)
 *   - xl/_rels/cellimages.xml.rels     rId → media 映射
 *   - [Content_Types].xml              + 各图片扩展名的 Default + cellimages.xml 的 Override
 *   - xl/_rels/workbook.xml.rels       + 指向 cellimages.xml 的 Relationship(WPS etCustomData 类型)
 *
 * 从模型重建(而非复制原件)→ 原文件已有的 + App 内新转的内嵌图一视同仁,导出后都在。
 */
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate'
import type { WorkbookModel } from '../model/types'

/**
 * WPS cellimages.xml 的内容类型 + workbook 关系类型。
 * ★ 逐字节对齐真·WPS 文件(2026-06 实测样本):内容类型是 **单数** cellimage,关系类型是 **2020/cellImage**。
 *   早期我用了复数 cellimages + 2017/etCustomData → WPS 加载不了登记表,DISPIMG 显示 #REF!。
 */
const CELLIMAGES_CT = 'application/vnd.wps-officedocument.cellimage+xml'
const CELLIMAGES_REL_TYPE = 'http://www.wps.cn/officeDocument/2020/cellImage'
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

/** 读 PNG 的像素尺寸 → EMU(供 cellimages.xml 的 <a:ext>;非 png/读不出时回落默认) */
function picExtEmu(bytes: Uint8Array | undefined, mime: string | undefined): { cx: number; cy: number } {
  const def = { cx: 990600, cy: 990600 }
  if (mime !== 'image/png' || !bytes || bytes.length < 24) return def
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (dv.getUint32(0) !== 0x89504e47) return def // PNG 签名
  const w = dv.getUint32(16)
  const h = dv.getUint32(20) // IHDR width/height(big-endian)
  if (!w || !h) return def
  return { cx: w * 9525, cy: h * 9525 } // 1px = 9525 EMU
}

const MIME_TO_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/svg+xml': 'svg',
  'image/webp': 'webp',
}

function xmlEscape(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

/** 写一个 zip 条目:复制成 ArrayBuffer-backed 的新 Uint8Array(对齐 fflate 的类型,绕开 ArrayBufferLike 摩擦) */
function put(files: Record<string, Uint8Array>, key: string, data: Uint8Array): void {
  const copy = new Uint8Array(data.byteLength)
  copy.set(data)
  files[key] = copy
}

/**
 * 把 WorkbookModel.cellImages 回注进 ExcelJS 写出的 zip 字节,返回新 zip 字节。
 * 无内嵌图(或全缺字节)→ 原样返回(零开销)。
 */
export function injectCellImagesIntoZip(zipBytes: Uint8Array, workbook: WorkbookModel): Uint8Array {
  const reg = workbook.cellImages
  if (!reg || reg.size === 0) return zipBytes

  // 只回注有字节的图(blob-only 的拿不到字节,跳过)
  const entries = Array.from(reg.values()).filter((ci) => ci.bytes && ci.bytes.length && ci.mime && MIME_TO_EXT[ci.mime])
  if (entries.length === 0) return zipBytes

  const files: Record<string, Uint8Array> = unzipSync(zipBytes)

  // 1. media + rels + cellimages.xml(逐条分配 rId / media 文件名)
  const relLines: string[] = []
  const picLines: string[] = []
  entries.forEach((ci, i) => {
    const n = i + 1
    const ext = MIME_TO_EXT[ci.mime!]
    const mediaName = `cellimage${n}.${ext}`
    put(files, `xl/media/${mediaName}`, ci.bytes!)
    const rId = `rId${n}`
    relLines.push(`<Relationship Id="${rId}" Type="${IMAGE_REL_TYPE}" Target="media/${mediaName}"/>`)
    const { cx, cy } = picExtEmu(ci.bytes, ci.mime)
    // 结构逐字节对齐真·WPS:cNvPr 带 descr、spPr 含 xfrm + prstGeom rect(空 spPr 会被 WPS 忽略 → 不显示图)
    picLines.push(
      `<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${n}" name="${xmlEscape(ci.id)}" descr="img"/><xdr:cNvPicPr/></xdr:nvPicPr>` +
        `<xdr:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
        `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>` +
        `</xdr:pic></etc:cellImage>`,
    )
  })

  put(
    files,
    'xl/cellimages.xml',
    strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<etc:cellImages xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
        ' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"' +
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
        picLines.join('') +
        '</etc:cellImages>',
    ),
  )
  put(
    files,
    'xl/_rels/cellimages.xml.rels',
    strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        relLines.join('') +
        '</Relationships>',
    ),
  )

  // 2. [Content_Types].xml:补各图片扩展名 Default + cellimages.xml Override
  const ctKey = '[Content_Types].xml'
  if (files[ctKey]) {
    let ct = strFromU8(files[ctKey])
    const usedExts = new Map<string, string>() // ext → mime
    for (const ci of entries) usedExts.set(MIME_TO_EXT[ci.mime!], ci.mime!)
    let inserts = ''
    for (const [ext, mime] of usedExts) {
      if (!new RegExp(`Extension="${ext}"`, 'i').test(ct)) inserts += `<Default Extension="${ext}" ContentType="${mime}"/>`
    }
    if (!ct.includes('/xl/cellimages.xml')) {
      inserts += `<Override PartName="/xl/cellimages.xml" ContentType="${CELLIMAGES_CT}"/>`
    }
    if (inserts) ct = ct.replace('</Types>', inserts + '</Types>')
    put(files, ctKey, strToU8(ct))
  }

  // 3. xl/_rels/workbook.xml.rels:加指向 cellimages.xml 的关系(rId 不与现有冲突)
  const wbRelsKey = 'xl/_rels/workbook.xml.rels'
  if (files[wbRelsKey]) {
    let rels = strFromU8(files[wbRelsKey])
    if (!rels.includes('cellimages.xml')) {
      let max = 0
      for (const m of rels.matchAll(/Id="rId(\d+)"/g)) max = Math.max(max, Number(m[1]))
      const rid = `rId${max + 1}`
      rels = rels.replace(
        '</Relationships>',
        `<Relationship Id="${rid}" Type="${CELLIMAGES_REL_TYPE}" Target="cellimages.xml"/></Relationships>`,
      )
      put(files, wbRelsKey, strToU8(rels))
    }
  }

  return zipSync(files)
}
