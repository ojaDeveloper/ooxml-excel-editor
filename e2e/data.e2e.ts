import { test, expect } from '@playwright/test'
import { loadSample } from './helpers'

test.describe('数据读取 API(组件 ref)', () => {
  test('getCellText / getSheetData / getSheetJSON 委托可用', async ({ page }) => {
    await loadSample(page)
    // 等数据就绪(getWorkbook 非空)再读
    await page.waitForFunction(() => (window as any).__excelViewer?.getWorkbook?.() != null, null, { timeout: 8_000 })

    // 示例表 0: 行0=合并标题,行1=表头(产品/单价/…),行2+=数据
    const headerA = await page.evaluate(() => (window as any).__excelViewer.getCellText(1, 0))
    expect(headerA).toBe('产品')

    // 整表 2D(默认显示文本)
    const grid10 = await page.evaluate(() => (window as any).__excelViewer.getSheetData()[1][0])
    expect(grid10).toBe('产品')

    // 原始值
    const raw = await page.evaluate(() => (window as any).__excelViewer.getSheetData({ format: false })[1][0])
    expect(raw).toBe('产品')

    // JSON(表头在第 1 行)→ 首条数据是第一个产品
    const json = await page.evaluate(() => (window as any).__excelViewer.getSheetJSON({ headerRow: 1 }))
    expect(Array.isArray(json)).toBeTruthy()
    expect(json[0]['产品']).toBe('笔记本电脑')
  })
})
