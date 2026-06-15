import { test, expect, type Page } from '@playwright/test'

/**
 * 批注编辑 e2e(1.11.0)。三壳共测 —— 批注对话框是框架无关 DOM(三壳共用一份)。
 * 编辑模式选格 → openCommentEditor → 填批注 → 确定 → getCellComment 读到 → undo 清除。
 */
const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 插入批注对话框 → 确定 → 读到 + undo 清除`, async ({ page }) => {
    await page.goto(url)
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => !!(window as any)[h]?.getWorkbook?.()?.sheets?.length, handle, { timeout: 20_000 })
    await page.locator('.edit-toggle').click()

    await call(page, handle, 'setSelection', { top: 3, left: 0, bottom: 3, right: 0 }) // A4(可编辑)
    await call(page, handle, 'openCommentEditor')
    const card = page.locator('.ooxml-comment-card')
    await expect(card).toBeVisible()
    await card.locator('[data-text]').fill('请复核此项')
    await card.locator('[data-ok]').click()
    await expect(card).toBeHidden()
    expect(await call(page, handle, 'getCellComment', 3, 0)).toBe('请复核此项')

    await call(page, handle, 'undo') // 批注单次撤销
    expect(await call(page, handle, 'getCellComment', 3, 0)).toBe('')
  })
}

test.describe('批注编辑 e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
  run('Vue2', 'http://localhost:5302/', 'canvas.ov-grid-canvas', '__excelViewerVue2')
})
