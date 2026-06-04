import { test, expect } from '@playwright/test'
import { loadSample } from './helpers'

test.describe('操作工具栏(可配置/可插件)', () => {
  test('默认显示 查找/筛选;查找开关 + 筛选 toggle 自动筛选', async ({ page }) => {
    await loadSample(page)
    const bar = page.locator('.action-toolbar')
    await expect(bar).toBeVisible()
    const findBtn = bar.getByRole('button', { name: /查找/ })
    const filterBtn = bar.getByRole('button', { name: /筛选/ })
    await expect(findBtn).toBeVisible()
    await expect(filterBtn).toBeVisible()

    // 查找按钮 → 开查找条;再点 → 关
    await findBtn.click()
    await expect(page.locator('.find-bar')).toBeVisible()
    await findBtn.click()
    await expect(page.locator('.find-bar')).toBeHidden()

    // 示例含自动筛选 → 筛选按钮 active;toggle 关 → active 消失;再 toggle → 回来
    await expect(filterBtn).toHaveClass(/active/)
    await filterBtn.click()
    await expect(filterBtn).not.toHaveClass(/active/)
    await filterBtn.click()
    await expect(filterBtn).toHaveClass(/active/)
  })
})
