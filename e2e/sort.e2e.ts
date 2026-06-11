import { test, expect, type Page } from '@playwright/test'

async function loadSample(page: Page) {
  await page.goto('/')
  await page.getByRole('button', { name: '加载示例' }).click()
  await expect(page.locator('canvas.grid-canvas')).toBeVisible({ timeout: 20_000 })
  await page.waitForFunction(
    () => {
      const c = document.querySelector('canvas.grid-canvas') as HTMLCanvasElement | null
      if (!c || !c.width) return false
      const ctx = c.getContext('2d')!
      const d = ctx.getImageData(0, 0, Math.min(c.width, 300), Math.min(c.height, 300)).data
      for (let i = 0; i < d.length; i += 4) if (d[i] < 248 || d[i + 1] < 248 || d[i + 2] < 248) return true
      return false
    },
    null,
    { timeout: 20_000 },
  )
  await page.waitForFunction(() => (window as any).__excelViewer?.rectOf?.(1, 0) != null, null, { timeout: 5_000 })
}

async function clickFilterButton(page: Page, headerRow: number, col: number) {
  const area = page.locator('.render-area')
  const box = (await area.boundingBox())!
  const rect = await page.evaluate(([r, c]) => (window as any).__excelViewer.rectOf(r, c), [headerRow, col] as const)
  await page.mouse.click(box.x + rect.x + rect.w - 8, box.y + rect.y + rect.h / 2)
}

// 用「单价」列(数值)排序,断言不依赖 locale(数值比较在哪都一致)。
// 销售报表 表头:[产品=0, 单价=1, 数量=2, 金额=3, 增长率=4]
const PRICE_COL = 1
test.describe('排序 e2e(筛选下拉 升序/降序 → 行重排)', () => {
  test('按单价升序→已升序;降序→反序;整行同步移动;数据集合不变', async ({ page }) => {
    await loadSample(page)
    // 自动筛选区 A2:E7 → 表头行 1,数据行 2..6
    const prices = () =>
      page.evaluate((c) => [2, 3, 4, 5, 6].map((r) => (window as any).__excelViewer.getCellValue(r, c) as number), PRICE_COL)
    const origSet = (await prices()).slice().sort((a, b) => a - b)

    // 升序
    await clickFilterButton(page, 1, PRICE_COL)
    await expect(page.locator('.filter-pop')).toBeVisible()
    await page.locator('.filter-pop .sort button').filter({ hasText: '升序' }).click()
    await expect(page.locator('.filter-pop')).toBeHidden()
    const asc = await prices()
    expect(asc).toEqual([...asc].sort((a, b) => a - b)) // 已升序
    expect([...asc].sort((a, b) => a - b)).toEqual(origSet) // 集合不变

    // 降序
    await clickFilterButton(page, 1, PRICE_COL)
    await expect(page.locator('.filter-pop')).toBeVisible()
    await page.locator('.filter-pop .sort button').filter({ hasText: '降序' }).click()
    await expect(page.locator('.filter-pop')).toBeHidden()
    const desc = await prices()
    expect(desc).toEqual([...asc].reverse()) // 正好反序

    // 整行随之移动:单价最高那行的"产品"同步到首数据行(笔记本电脑 ¥5999 最贵 → 降序后在 r2)
    const topProduct = await page.evaluate(() => (window as any).__excelViewer.getCellText(2, 0))
    expect(topProduct).toBe('笔记本电脑')
  })

  test('工具栏排序按活动列升序', async ({ page }) => {
    await loadSample(page)
    await page.evaluate((c) => (window as any).__excelViewer.setSelection({ top: 2, left: c, bottom: 2, right: c }), PRICE_COL)
    await page.locator('.action-toolbar').getByRole('button', { name: /排序/ }).click()
    await page.locator('.tb-menu').getByRole('button', { name: /升序/ }).click()

    const asc = await page.evaluate((c) => [2, 3, 4, 5, 6].map((r) => (window as any).__excelViewer.getCellValue(r, c) as number), PRICE_COL)
    expect(asc).toEqual([...asc].sort((a, b) => a - b))
  })
})
