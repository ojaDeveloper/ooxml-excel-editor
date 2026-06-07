/** 生成一个简单的渲染模板 .xlsx —— 用于 demo 验证 P3 模板填值。输出到 public/template-sample.xlsx。
 *  样式:醒目标题 + 表头 + 占位符 {{customer}}/{{total}}/{{date}} + A5 起的明细表锚点(模板自带 5 行空表头) */
import ExcelJS from 'exceljs'
import { mkdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const __dirname = dirname(fileURLToPath(import.meta.url))
const outDir = join(__dirname, '..', 'public')
mkdirSync(outDir, { recursive: true })

const wb = new ExcelJS.Workbook()
const s = wb.addWorksheet('发票模板', { views: [{ state: 'frozen', xSplit: 0, ySplit: 4 }] })

// 表头区(行 1-3,大字 + 醒目背景)
s.mergeCells('A1:E1')
s.getCell('A1').value = '订单结算单'
s.getCell('A1').font = { bold: true, size: 20, color: { argb: 'FFFFFFFF' } }
s.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }
s.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF21A366' } }
s.getRow(1).height = 36

// 占位符段(行 2-3,模板 demo 填值用)
s.getCell('A2').value = '客户:'
s.getCell('A2').font = { bold: true }
s.getCell('B2').value = '{{customer}}'
s.mergeCells('B2:E2')
s.getCell('A3').value = '日期:'
s.getCell('A3').font = { bold: true }
s.getCell('B3').value = '{{date}}'

// 明细表头(行 4,锚点表的标题行,JSON 数据从 A5 起填)
const headers = ['商品', '单价', '数量', '金额', '备注']
s.getRow(4).values = headers
s.getRow(4).eachCell((cell) => {
  cell.font = { bold: true }
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } }
  cell.border = {
    top: { style: 'thin' }, bottom: { style: 'medium' },
    left: { style: 'thin' }, right: { style: 'thin' },
  }
  cell.alignment = { horizontal: 'center' }
})

// 给明细区(A5:E15)预设边框,这样 JSON 填上去就有格子边框
for (let r = 5; r <= 15; r++) {
  for (let c = 1; c <= 5; c++) {
    s.getCell(r, c).border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } }
  }
}

// 合计行(行 16)
s.getCell('A16').value = '合计:'
s.getCell('A16').font = { bold: true }
s.mergeCells('B16:C16')
s.getCell('D16').value = '{{total}}'
s.getCell('D16').font = { bold: true, color: { argb: 'FF21A366' } }
s.getCell('D16').alignment = { horizontal: 'right' }

// 列宽
s.columns = [
  { width: 22 }, { width: 12 }, { width: 10 }, { width: 14 }, { width: 18 },
]

await wb.xlsx.writeFile(join(outDir, 'template-sample.xlsx'))
console.log('✓ 生成 public/template-sample.xlsx(发票模板,含 {{customer}}/{{date}}/{{total}} 占位符 + A5 起的明细表锚点)')
