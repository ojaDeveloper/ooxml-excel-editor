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
  // 等命令式 API 挂到 window 且能算出几何
  await page.waitForFunction(() => (window as any).__excelViewer?.rectOf?.(1, 0) != null, null, { timeout: 5_000 })
}

const spacerH = (page: Page) =>
  page.evaluate(() => (document.querySelector('.scroller .spacer') as HTMLElement).offsetHeight)

/** 点击某列(自动筛选表头行)的下拉按钮 */
async function clickFilterButton(page: Page, headerRow: number, col: number) {
  const area = page.locator('.render-area')
  const box = (await area.boundingBox())!
  const rect = await page.evaluate(
    ([r, c]) => (window as any).__excelViewer.rectOf(r, c),
    [headerRow, col] as const,
  )
  // 按钮在格子右侧 16px 方块 → 点其中心
  await page.mouse.click(box.x + rect.x + rect.w - 8, box.y + rect.y + rect.h / 2)
}

test.describe('筛选 e2e(自动筛选下拉 → 隐藏行)', () => {
  test('开筛选浮层 → 取消勾选某值隐藏行 → 清除恢复', async ({ page }) => {
    await loadSample(page)
    // 表头行 A2:E7 → 0-based 行 1;'产品' 列 = 0
    const before = await spacerH(page)

    await clickFilterButton(page, 1, 0)
    const pop = page.locator('.filter-pop')
    await expect(pop).toBeVisible()

    // 列出了去重值(5 个产品)
    const rowCount = await pop.locator('.list .row').count()
    expect(rowCount).toBeGreaterThanOrEqual(2)

    // 取消勾选第一个值 → 确定 → 该行被隐藏(spacer 变矮)
    await pop.locator('.list .row input[type=checkbox]').first().uncheck()
    await pop.getByRole('button', { name: '确定' }).click()
    await expect(pop).toBeHidden()
    await expect.poll(() => spacerH(page)).toBeLessThan(before)

    // 重新打开 → 清除筛选 → 恢复原高度
    await clickFilterButton(page, 1, 0)
    await expect(page.locator('.filter-pop')).toBeVisible()
    await page.getByRole('button', { name: '清除筛选' }).click()
    await expect.poll(() => spacerH(page)).toBe(before)
  })
})
