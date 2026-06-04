import { expect, type Page } from '@playwright/test'

/** 加载示例并等待真实渲染(canvas 像素非空) */
export async function loadSample(page: Page) {
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
