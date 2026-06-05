/**
 * 生成一个含 WPS 单元格内嵌图(DISPIMG)的 .xlsx,用于解析/渲染验证。输出 public/wps-dispimg-sample.xlsx。
 *
 * 做法:ExcelJS 先写一个普通工作簿(B2 放 DISPIMG 公式),再用 fflate 把 WPS 私有件
 * xl/cellimages.xml + 其 rels + media 注入到 zip 里 —— 这正是 WPS 导出文件的结构。
 */
import ExcelJS from 'exceljs'
import { unzipSync, zipSync, strToU8 } from 'fflate'
import { mkdirSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const __dirname = dirname(fileURLToPath(import.meta.url))
const outDir = join(__dirname, '..', 'public')
mkdirSync(outDir, { recursive: true })

const wb = new ExcelJS.Workbook()
const ws = wb.addWorksheet('内嵌图')
ws.getCell('A1').value = 'WPS 单元格内嵌图(DISPIMG)演示'
ws.getCell('A1').font = { bold: true, size: 14 }
ws.getCell('A2').value = '下面这格的图是嵌在单元格里的(非浮动):'
ws.getCell('B2').value = { formula: '_xlfn.DISPIMG("ID_demo_0001",1)', result: ' ' }
ws.getColumn('B').width = 16
ws.getRow(2).height = 80

const baseBuf = await wb.xlsx.writeBuffer()
const zip = unzipSync(new Uint8Array(baseBuf))

// 1x1 透明 png(占位图;真实 WPS 会是任意图片)
const PNG_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const pngBytes = Uint8Array.from(Buffer.from(PNG_B64, 'base64'))

const cellImagesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="ID_demo_0001"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr/></xdr:pic></etc:cellImage>
</etc:cellImages>`

const cellImagesRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/dispimg1.png"/>
</Relationships>`

zip['xl/cellimages.xml'] = strToU8(cellImagesXml)
zip['xl/_rels/cellimages.xml.rels'] = strToU8(cellImagesRels)
zip['xl/media/dispimg1.png'] = pngBytes

const out = zipSync(zip)
const outPath = join(outDir, 'wps-dispimg-sample.xlsx')
writeFileSync(outPath, out)
console.log('已生成', outPath, `(${out.length} bytes)`)
