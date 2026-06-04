import { test, expect } from '@playwright/test'
import { loadSample } from './helpers'

test.describe('操作工具栏(可配置/可插件/响应式)', () => {
  test('查找开关 + 筛选 toggle 自动筛选', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await loadSample(page)
    const bar = page.locator('.action-toolbar')
    await expect(bar).toBeVisible()
    const findBtn = bar.getByRole('button', { name: /查找/ })
    const filterBtn = bar.getByRole('button', { name: /筛选/, exact: false }).first()

    await findBtn.click()
    await expect(page.locator('.find-bar')).toBeVisible()
    await findBtn.click()
    await expect(page.locator('.find-bar')).toBeHidden()

    await expect(filterBtn).toHaveClass(/active/) // 示例含 autofilter
    await filterBtn.click()
    await expect(filterBtn).not.toHaveClass(/active/)
    await filterBtn.click()
    await expect(filterBtn).toHaveClass(/active/)
  })

  test('清除筛选 初始禁用(无筛选时)', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await loadSample(page)
    const btn = page.locator('.action-toolbar').getByRole('button', { name: /清除筛选/ })
    await expect(btn).toBeDisabled()
  })

  test('导出下拉(工具栏)→ 触发 PNG 下载', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 })
    await loadSample(page)
    const bar = page.locator('.action-toolbar')
    await bar.getByRole('button', { name: /导出/ }).first().click()
    const menu = page.locator('.tb-menu')
    await expect(menu).toBeVisible()
    const dl = page.waitForEvent('download', { timeout: 30_000 })
    await menu.getByRole('button', { name: '导出为图片 (PNG)' }).click()
    expect((await dl).suggestedFilename()).toMatch(/\.png$/)
  })

  test('响应式溢出: 窄屏出现「更多」并能展开,宽屏收起', async ({ page }) => {
    await loadSample(page)
    const bar = page.locator('.action-toolbar')
    await page.setViewportSize({ width: 1280, height: 800 })
    await expect(bar.locator('.more')).toHaveCount(0)

    await page.setViewportSize({ width: 400, height: 800 })
    await expect(bar.locator('.more')).toBeVisible()
    await bar.locator('.more .tool').click()
    await expect(page.locator('.tb-menu')).toBeVisible()

    await page.setViewportSize({ width: 1280, height: 800 })
    await expect(bar.locator('.more')).toHaveCount(0)
  })
})
