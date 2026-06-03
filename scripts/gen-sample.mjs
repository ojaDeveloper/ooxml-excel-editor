/** 生成一个特性丰富的示例 .xlsx 用于预览验证。输出到 public/sample.xlsx。 */
import ExcelJS from 'exceljs'
import { mkdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const __dirname = dirname(fileURLToPath(import.meta.url))
const outDir = join(__dirname, '..', 'public')
mkdirSync(outDir, { recursive: true })

const wb = new ExcelJS.Workbook()

// ---- Sheet 1: 销售报表(样式/数字格式/合并/边框/冻结/筛选) ----
const s1 = wb.addWorksheet('销售报表', {
  views: [{ state: 'frozen', xSplit: 1, ySplit: 2 }],
})
s1.mergeCells('A1:E1')
s1.getCell('A1').value = '2026 年度销售汇总'
s1.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } }
s1.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }
s1.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF21A366' } }
s1.getRow(1).height = 30

const headers = ['产品', '单价', '数量', '金额', '增长率']
s1.getRow(2).values = headers
s1.getRow(2).eachCell((cell) => {
  cell.font = { bold: true }
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } }
  cell.border = {
    top: { style: 'thin' }, bottom: { style: 'medium' },
    left: { style: 'thin' }, right: { style: 'thin' },
  }
  cell.alignment = { horizontal: 'center' }
})

const rows = [
  ['笔记本电脑', 5999, 120, null, 0.152],
  ['机械键盘', 399, 530, null, -0.043],
  ['显示器', 1299, 210, null, 0.087],
  ['鼠标', 89, 1200, null, 0.231],
  ['耳机', 599, 340, null, -0.012],
]
rows.forEach((r, i) => {
  const rowNum = i + 3
  const row = s1.getRow(rowNum)
  row.getCell(1).value = r[0]
  row.getCell(2).value = r[1]
  row.getCell(2).numFmt = '¥#,##0.00'
  row.getCell(3).value = r[2]
  row.getCell(3).numFmt = '#,##0'
  row.getCell(4).value = { formula: `B${rowNum}*C${rowNum}`, result: r[1] * r[2] }
  row.getCell(4).numFmt = '¥#,##0;[Red]-¥#,##0'
  row.getCell(5).value = r[4]
  row.getCell(5).numFmt = '0.0%;[Red]-0.0%'
  row.eachCell((cell) => {
    cell.border = { top: { style: 'hair' }, bottom: { style: 'hair' }, left: { style: 'thin' }, right: { style: 'thin' } }
  })
})
s1.getColumn(1).width = 16
s1.getColumn(2).width = 14
s1.getColumn(3).width = 10
s1.getColumn(4).width = 16
s1.getColumn(5).width = 12
s1.autoFilter = 'A2:E2'

// 数据条 + 色阶条件格式
s1.addConditionalFormatting({
  ref: 'C3:C7',
  rules: [{ type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' } }],
})
s1.addConditionalFormatting({
  ref: 'E3:E7',
  rules: [{
    type: 'colorScale',
    cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
    color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }],
  }],
})

// ---- Sheet 2: 日期与格式样例 ----
const s2 = wb.addWorksheet('格式样例')
const samples = [
  ['日期(年月日)', new Date(Date.UTC(2026, 0, 15)), 'yyyy"年"m"月"d"日"'],
  ['日期(短)', new Date(Date.UTC(2026, 5, 2)), 'yyyy-mm-dd'],
  ['千分位', 1234567.891, '#,##0.00'],
  ['百分比', 0.8734, '0.00%'],
  ['货币', -8800, '¥#,##0.00;[Red]-¥#,##0.00'],
  ['科学计数', 0.00001234, '0.00E+00'],
  ['文本', 'Hello 世界', '@'],
]
s2.getColumn(1).width = 16
s2.getColumn(2).width = 22
s2.getColumn(3).width = 26
s2.getRow(1).values = ['说明', '值', '格式代码']
s2.getRow(1).font = { bold: true }
samples.forEach((sp, i) => {
  const row = s2.getRow(i + 2)
  row.getCell(1).value = sp[0]
  row.getCell(2).value = sp[1]
  row.getCell(2).numFmt = sp[2]
  row.getCell(3).value = sp[2]
  row.getCell(3).font = { name: 'Consolas', color: { argb: 'FF888888' } }
})

// 保真/交互演示区(文本溢出 / 批注 / 数据验证 / 超链接)
const demoRow = samples.length + 4
s2.getCell(`A${demoRow}`).value = '保真演示'
s2.getCell(`A${demoRow}`).font = { bold: true }

// 文本溢出: 长文本 + 右侧空格
s2.getCell(`A${demoRow + 1}`).value = '这是一段很长很长的文本用来演示溢出到右侧空白单元格的效果'
// B/C 列留空，文字应溢出铺过去

// 批注
const noteCell = s2.getCell(`A${demoRow + 2}`)
noteCell.value = '带批注的格(右上角红三角)'
noteCell.note = '这是一条单元格批注\n鼠标悬停可看到全文。'

// 数据验证(列表)
const dvCell = s2.getCell(`A${demoRow + 3}`)
dvCell.value = '选我看下拉箭头'
dvCell.dataValidation = {
  type: 'list',
  allowBlank: true,
  formulae: ['"待发货,已发货,已签收"'],
}

// 超链接
const linkCell = s2.getCell(`A${demoRow + 4}`)
linkCell.value = { text: '点我打开链接', hyperlink: 'https://example.com' }
linkCell.font = { color: { argb: 'FF0563C1' }, underline: true }

// ---- Sheet 3: 大表(性能) ----
const s3 = wb.addWorksheet('大表')
for (let r = 1; r <= 2000; r++) {
  s3.getRow(r).values = [r, `行${r}`, Math.round(Math.random() * 10000) / 100, r % 2 === 0]
}

const outPath = join(outDir, 'sample.xlsx')
await wb.xlsx.writeFile(outPath)
console.log('已生成:', outPath)
