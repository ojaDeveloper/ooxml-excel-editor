/**
 * 纯 Node 高保真往返示例 —— 打开真实 .xlsx → 程序化改值/样式 → 保样式回写 .xlsx。
 *
 * 先构建产物:  npm run build
 * 再运行:      node examples/node-roundtrip.mjs
 *
 * overlay 模式重载原件叠加编辑,保留原件的样式/条件格式/图片/透视表(裸 ExcelJS 会丢)。
 * 真实项目里改成 import ... from 'ooxml-excel-editor/core'(装包后)。
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  openWorkbook,
  parseWorkbook,
  setCellValue,
  applyStyleOverride,
  workbookToXlsxBytes,
  getCellText,
} from '../dist/core.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const SAMPLE = join(__dirname, '..', 'public', 'sample.xlsx')
const OUT = join(__dirname, 'out.xlsx')

const src = readFileSync(SAMPLE)
const wb = await openWorkbook(src)
const sheet = wb.sheets[0]

// 改一个空的远端格 + 给它加粗红字(避开合并标题区)
setCellValue(sheet, 100, 0, 'Node 改的')
applyStyleOverride(sheet, 100, 0, { font: { bold: true, color: '#FF0000' } })

// overlay 高保真回写;返回 Uint8Array,直接 fs 落盘(无需 Blob)
const bytes = await workbookToXlsxBytes(wb, {
  fidelity: 'overlay',
  sourceBuffer: src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength),
})
writeFileSync(OUT, bytes)
console.log(`已写出 ${OUT}（${bytes.byteLength} 字节）`)

// 回读校验
const wb2 = await parseWorkbook(readFileSync(OUT))
console.log('回读 (100,0) =', getCellText(wb2.sheets[0], 100, 0))
