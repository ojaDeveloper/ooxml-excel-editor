/**
 * 纯 Node 取数示例 —— 解析 .xlsx → 拿"人看到的"数据(显示文本/JSON/CSV)。
 *
 * 先构建产物:  npm run build
 * 再运行:      node examples/node-extract.mjs
 *
 * 真实项目里改成 import ... from 'ooxml-excel-editor/core'(装包后)。
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  openWorkbook,
  getSheetData,
  sheetToJSON,
  getCellText,
  toCsv,
} from '../dist/core.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const SAMPLE = join(__dirname, '..', 'public', 'sample.xlsx')

// openWorkbook 直接吃 Node Buffer —— 不用手动转 ArrayBuffer
const wb = await openWorkbook(readFileSync(SAMPLE))
const sheet = wb.sheets[0]

console.log('工作表:', wb.sheets.map((s) => s.name).join(' / '))
console.log('\n— 显示文本 A1 —')
console.log(getCellText(sheet, 0, 0))

console.log('\n— getSheetData(前 3 行, 显示文本) —')
console.log(getSheetData(sheet, { format: true }).slice(0, 3))

console.log('\n— sheetToJSON(前 2 条) —')
console.log(sheetToJSON(sheet).slice(0, 2))

console.log('\n— toCsv(前 200 字) —')
console.log(toCsv(sheet).slice(0, 200))
