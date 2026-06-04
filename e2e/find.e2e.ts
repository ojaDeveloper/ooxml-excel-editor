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
}

test.describe('查找 e2e(Ctrl+F · 高亮 · 计数 · 导航)', () => {
  test('Ctrl+F 打开 → 命中计数 / 无结果 / 上下导航 / Esc 关闭', async ({ page }) => {
    await loadSample(page)
    await page.locator('.scroller').click() // 聚焦网格,Ctrl+F 才会被根容器捕获
    await page.keyboard.press('Control+f')

    const bar = page.locator('.find-bar')
    await expect(bar).toBeVisible()

    // 存在的词(列头)→ "x/N"
    await bar.locator('input.q').fill('产品')
    await expect(bar.locator('.count')).toHaveText(/^\d+\/\d+$/)

    // 不存在的词 → 无结果
    await bar.locator('input.q').fill('zz不存在的词zz')
    await expect(bar.locator('.count')).toHaveText('无结果')

    // 多命中词("9" 命中多个单价单元格)→ 下一个推进当前项
    await bar.locator('input.q').fill('9')
    const txt = await bar.locator('.count').innerText()
    const m = txt.match(/^(\d+)\/(\d+)$/)
    expect(m).not.toBeNull()
    const total = Number(m![2])
    expect(total).toBeGreaterThan(0)
    if (total >= 2) {
      await page.getByTitle('下一个 (Enter)').click()
      await expect(bar.locator('.count')).toHaveText(`2/${total}`)
      await page.getByTitle('上一个 (Shift+Enter)').click()
      await expect(bar.locator('.count')).toHaveText(`1/${total}`)
    }

    // Esc 关闭
    await bar.locator('input.q').press('Escape')
    await expect(bar).toBeHidden()
  })
})
