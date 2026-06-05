/**
 * 解析 WPS 私有的"单元格内嵌图"(DISPIMG)。
 *
 * WPS 把图放进单元格时不走标准 drawing 锚点,而是:
 *   1. 单元格存一条公式 `=_xlfn.DISPIMG("ID_xxx",1)`(我们的模型里落在 CellModel.formula)
 *   2. 图片登记在 workbook 级私有件 `xl/cellimages.xml`:每个 <etc:cellImage> 包一个 <xdr:pic>,
 *      其 cNvPr@name = DISPIMG id,blip@embed → rels → media/imageN.png
 *   3. `xl/_rels/cellimages.xml.rels` 给出 rId → 真实图片路径
 *
 * 标准解析器(drawing-parser)看不到这种图 → 这类 WPS 文件打开就缺图。本模块补齐:
 *   - 读 cellimages.xml 建 id → CellImage(bytes/mime)登记表(workbook 级)
 *   - 扫各表单元格公式,命中 DISPIMG 的把 id 记到 CellModel.dispImgId(渲染层据此画图)
 */
import type { RawPackage } from './raw-xml'
import { parseRels, toArray, basename } from './raw-xml'
import type { CellImage, SheetModel } from '../model/types'

const MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
}

/** 从单元格公式里抽 DISPIMG id;非 DISPIMG 返 undefined */
export function dispImgIdOf(formula: string | undefined): string | undefined {
  if (!formula || formula.indexOf('DISPIMG') < 0) return undefined
  const m = /DISPIMG\s*\(\s*"([^"]+)"/i.exec(formula)
  return m ? m[1] : undefined
}

/**
 * 解析 xl/cellimages.xml → id→CellImage 登记表,并回填各表单元格的 dispImgId。
 * 无该私有件(非 WPS 文件)→ 返 undefined,不动任何单元格。
 */
export function attachCellImages(pkg: RawPackage, sheets: SheetModel[]): Map<string, CellImage> | undefined {
  const registry = parseRegistry(pkg)

  // 不管有没有登记表,都扫一遍公式标记 dispImgId(有 id 没图时渲染层画占位)
  let anyMarked = false
  for (const sheet of sheets) {
    for (const cell of sheet.cells.values()) {
      if (cell.type !== 'formula') continue
      const id = dispImgIdOf(cell.formula)
      if (id) {
        cell.dispImgId = id
        anyMarked = true
      }
    }
  }

  if (!registry && !anyMarked) return undefined
  return registry ?? new Map()
}

/** 读 cellimages.xml + 其 rels → id→CellImage(bytes/mime);无私有件返 undefined */
function parseRegistry(pkg: RawPackage): Map<string, CellImage> | undefined {
  const xml = pkg.parse('xl/cellimages.xml')
  const root = xml?.cellImages
  if (!root) return undefined
  const rels = parseRels(pkg, 'xl/cellimages.xml')
  const out = new Map<string, CellImage>()

  for (const ci of toArray(root.cellImage)) {
    const pic = ci?.pic
    if (!pic) continue
    const id = pic.nvPicPr?.cNvPr?.['@_name']
    const embed = pic.blipFill?.blip?.['@_embed'] ?? pic.blipFill?.blip?.['@_link']
    if (!id || !embed) continue
    const target = rels[embed]
    if (!target) continue
    const data = imageData(pkg, target)
    if (!data) continue
    out.set(String(id), { id: String(id), bytes: data.bytes, mime: data.mime, src: '' })
  }

  return out.size ? out : undefined
}

function imageData(pkg: RawPackage, path: string): { bytes: Uint8Array; mime: string } | undefined {
  const bytes = pkg.bytes(path)
  if (!bytes) return undefined
  const ext = basename(path).split('.').pop()?.toLowerCase() || 'png'
  const mime = MIME[ext]
  if (!mime) return undefined // emf/wmf 等浏览器不支持,跳过
  return { bytes, mime }
}
