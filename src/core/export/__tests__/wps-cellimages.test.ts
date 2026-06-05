import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { unzipSync, strFromU8 } from 'fflate'
import { parseWorkbook } from '../../parser/index'
import { workbookToXlsxBlob } from '../xlsx-writer'
import { injectCellImagesIntoZip } from '../wps-cellimages'
import { cellKey } from '../../model/types'
import type { WorkbookModel } from '../../model/types'

function loadWps(): ArrayBuffer {
  const buf = readFileSync(join(__dirname, '..', '..', '..', '..', 'public', 'wps-dispimg-sample.xlsx'))
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)
}

async function reparse(blob: Blob): Promise<WorkbookModel> {
  const ab = await blob.arrayBuffer()
  return parseWorkbook(ab)
}

// 第三期:导出回注 WPS 内嵌图(DISPIMG)→ 解析→导出→再解析 往返存活
describe('WPS 内嵌图导出回注(第三期:DISPIMG 往返 .xlsx)', () => {
  it('rebuild 模式:导出后再解析,cellImages + dispImgId 仍在', async () => {
    const model = await parseWorkbook(loadWps())
    expect(model.cellImages!.has('ID_demo_0001')).toBe(true)

    const blob = await workbookToXlsxBlob(model, { fidelity: 'rebuild' })
    const back = await reparse(blob)
    expect(back.cellImages).toBeDefined()
    expect(back.cellImages!.size).toBe(1)
    const ci = back.cellImages!.get('ID_demo_0001')!
    expect(ci.mime).toBe('image/png')
    expect(ci.bytes!.length).toBeGreaterThan(0)
    expect(back.sheets[0].cells.get(cellKey(1, 1))?.dispImgId).toBe('ID_demo_0001')
  })

  it('overlay 模式:导出后再解析,cellImages 仍在', async () => {
    const src = loadWps()
    const model = await parseWorkbook(src)
    const blob = await workbookToXlsxBlob(model, { fidelity: 'overlay', sourceBuffer: src })
    const back = await reparse(blob)
    expect(back.cellImages!.has('ID_demo_0001')).toBe(true)
  })

  it('回注的 zip 含全套零件 + Content_Types/workbook-rels 已打补丁', async () => {
    const model = await parseWorkbook(loadWps())
    const blob = await workbookToXlsxBlob(model, { fidelity: 'rebuild' })
    const z = unzipSync(new Uint8Array(await blob.arrayBuffer()))
    // 私有件齐全
    expect(z['xl/cellimages.xml']).toBeDefined()
    expect(z['xl/_rels/cellimages.xml.rels']).toBeDefined()
    expect(Object.keys(z).some((k) => k.startsWith('xl/media/cellimage'))).toBe(true)
    // Content_Types:png Default + cellimage(单数)Override —— 逐字节对齐真·WPS
    const ct = strFromU8(z['[Content_Types].xml'])
    expect(ct).toContain('Extension="png"')
    expect(ct).toContain('PartName="/xl/cellimages.xml" ContentType="application/vnd.wps-officedocument.cellimage+xml"')
    // workbook.xml.rels:关系类型 2020/cellImage(非 2017/etCustomData)
    const wbRels = strFromU8(z['xl/_rels/workbook.xml.rels'])
    expect(wbRels).toContain('Type="http://www.wps.cn/officeDocument/2020/cellImage"')
    expect(wbRels).toContain('Target="cellimages.xml"')
    // cellimages.xml:cNvPr name=id + descr,spPr 含 xfrm/prstGeom(非空)
    const cimg = strFromU8(z['xl/cellimages.xml'])
    expect(cimg).toContain('name="ID_demo_0001"')
    expect(cimg).toContain('descr=')
    expect(cimg).toMatch(/<xdr:spPr><a:xfrm>.*<a:prstGeom prst="rect">/)
    // DISPIMG 单元格缓存值 <v>=DISPIMG("id",1)</v>
    const sheetKey = Object.keys(z).find((k) => /xl\/worksheets\/sheet1\.xml$/.test(k))!
    expect(strFromU8(z[sheetKey])).toContain('=DISPIMG(&quot;ID_demo_0001&quot;,1)</v>')
  })

  it('App 内浮动→嵌入转换后导出:新内嵌图也往返存活', async () => {
    const { addImage, convertFloatToCellImage } = await import('../../model/mutations')
    // 拿一个普通(非 WPS)工作簿,加一张带字节的浮动图,转成内嵌图
    const sampleBuf = readFileSync(join(__dirname, '..', '..', '..', '..', 'public', 'sample.xlsx'))
    const model = await parseWorkbook(sampleBuf.buffer.slice(sampleBuf.byteOffset, sampleBuf.byteOffset + sampleBuf.byteLength))
    const PNG = Uint8Array.from(
      Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64'),
    )
    const sheet = model.sheets[0]
    addImage(sheet, { src: '', bytes: PNG, mime: 'image/png', from: { col: 0, row: 0, colOffEmu: 0, rowOffEmu: 0 } })
    const id = convertFloatToCellImage(model, sheet, sheet.images.length - 1, 8, 8)
    expect(id).toBeTruthy()
    expect(model.cellImages!.size).toBe(1)

    const blob = await workbookToXlsxBlob(model, { fidelity: 'rebuild' })
    const back = await reparse(blob)
    expect(back.cellImages!.size).toBe(1)
    expect(back.sheets[0].cells.get(cellKey(8, 8))?.dispImgId).toBe(id)
  })

  it('无内嵌图工作簿:injectCellImagesIntoZip 原样返回(零开销)', () => {
    const fake = new Uint8Array([1, 2, 3])
    const wb = { sheets: [], cellImages: undefined } as unknown as WorkbookModel
    expect(injectCellImagesIntoZip(fake, wb)).toBe(fake)
  })
})
