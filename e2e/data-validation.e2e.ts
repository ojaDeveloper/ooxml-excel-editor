import { test, expect, type Page } from '@playwright/test'
import ExcelJS from 'exceljs'
import { writeFileSync, rmSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE = join(__dirname, '..', 'public', '_dv.xlsx')

test.beforeAll(async () => {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('S')
  ws.getCell('A1').value = '产品'
  ws.getCell('C6').value = '苹果'
  ws.getCell('C6').dataValidation = { type: 'list', allowBlank: true, formulae: ['"苹果,香蕉,橙子"'] }
  const buf = await wb.xlsx.writeBuffer()
  writeFileSync(FIXTURE, Buffer.from(buf as ArrayBuffer))
})
test.afterAll(() => { try { rmSync(FIXTURE) } catch { /* ignore */ } })

const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

function run(label: string, url: string, canvasSel: string, renderAreaSel: string, handle: string) {
  test(`${label}: 数据验证下拉箭头 → 弹选项菜单 → 点选填值(undo 可回退)`, async ({ page }) => {
    await page.goto(url)
    // 经 demo 文件选择器加载(让 demo 的 src=本文件,避免 :src prop 把示例重载回来覆盖)
    await page.locator('input[type="file"]').first().setInputFiles(FIXTURE)
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => { const v = (window as any)[h]; const s = v?.getWorkbook()?.sheets?.[v.getActiveSheet()]; return !!s?.dataValidationLists?.length }, handle, { timeout: 20_000 })
    await page.locator('.edit-toggle').click() // 进编辑模式(Vue=button / React=label+checkbox 都带 .edit-toggle)

    // 选中 C6(行5列2)→ 下拉箭头画在该格右上
    await call(page, handle, 'setSelection', { top: 5, left: 2, bottom: 5, right: 2 })
    const box = (await page.locator(renderAreaSel).boundingBox())!
    const rect = (await call(page, handle, 'rectOf', 5, 2)) as { x: number; y: number; w: number; h: number }
    // 点格内右侧、垂直居中的下拉按钮(filterButtonBox 把按钮放右侧居中)
    await page.mouse.click(box.x + rect.x + rect.w - 8, box.y + rect.y + rect.h / 2)

    const menu = page.locator('.ooxml-context-menu')
    await expect(menu).toBeVisible()
    await expect(menu).toContainText('香蕉')
    await menu.getByText('香蕉').click()
    expect(await call(page, handle, 'getCellValue', 5, 2)).toBe('香蕉') // 点选填进去了
    await call(page, handle, 'undo')
    expect(await call(page, handle, 'getCellValue', 5, 2)).toBe('苹果') // 可撤销
  })
}

test.describe('数据验证列表下拉 e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.render-area', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-render-area', '__excelViewerReact')
})
