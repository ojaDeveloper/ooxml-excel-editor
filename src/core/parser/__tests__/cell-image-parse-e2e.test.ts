import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { parseWorkbook } from '../index'
import { finalizeImages } from '../../finalize'
import { cellKey } from '../../model/types'

function loadWpsSample(): ArrayBuffer {
  const buf = readFileSync(join(__dirname, '..', '..', '..', '..', 'public', 'wps-dispimg-sample.xlsx'))
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)
}

// 端到端:真实生成的 WPS DISPIMG 文件 → parseWorkbook 应建登记表 + 标记单元格 dispImgId
describe('WPS DISPIMG 端到端解析', () => {
  it('cellImages 登记表 + B2 单元格 dispImgId', async () => {
    const wb = await parseWorkbook(loadWpsSample())
    expect(wb.cellImages).toBeDefined()
    expect(wb.cellImages!.has('ID_demo_0001')).toBe(true)
    const ci = wb.cellImages!.get('ID_demo_0001')!
    expect(ci.mime).toBe('image/png')
    expect(ci.bytes && ci.bytes.length).toBeGreaterThan(0)

    // B2 = (row 1, col 1) 0-based,公式 DISPIMG → dispImgId 已标记
    const b2 = wb.sheets[0].cells.get(cellKey(1, 1))
    expect(b2?.dispImgId).toBe('ID_demo_0001')
  })

  it('finalizeImages 给登记表落 blob/ data url(node 无 URL 时跳过不报错)', async () => {
    const wb = await parseWorkbook(loadWpsSample())
    // node 环境通常无 URL.createObjectURL → finalize 应安全跳过(src 仍空),不抛错
    expect(() => finalizeImages(wb)).not.toThrow()
  })
})
