import { test, expect, type Page } from '@playwright/test'

/**
 * 数字格式编辑 e2e(1.11.0)。三壳共测 —— 对话框是框架无关 DOM(三壳共用一份)。
 *  - API:setSelectionNumberFormat('0%') 把 0.5 显示成 50% + undo
 *  - 对话框:工具栏「数字格式」→ 改格式代码 → 确定 → 显示生效
 */
const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

function run(label: string, url: string, canvasSel: string, toolbarSel: string, handle: string) {
  test(`${label}: API setSelectionNumberFormat 0% → 0.5 显示 50% + undo`, async ({ page }) => {
    await page.goto(url)
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => !!(window as any)[h]?.getWorkbook?.()?.sheets?.length, handle, { timeout: 20_000 })
    await page.locator('.edit-toggle').click()

    await call(page, handle, 'editCell', 3, 0, '0.5') // A4 = 0.5(避开只读行）
    await call(page, handle, 'setSelection', { top: 3, left: 0, bottom: 3, right: 0 })
    await call(page, handle, 'setSelectionNumberFormat', '0%')
    expect(await call(page, handle, 'getCellText', 3, 0)).toBe('50%')
    await call(page, handle, 'undo')
    expect(await call(page, handle, 'getCellText', 3, 0)).not.toBe('50%')
  })

  test(`${label}: 数字格式对话框 改格式代码 → 确定生效`, async ({ page }) => {
    await page.goto(url)
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => !!(window as any)[h]?.getWorkbook?.()?.sheets?.length, handle, { timeout: 20_000 })
    await page.locator('.edit-toggle').click()

    await call(page, handle, 'editCell', 3, 0, '1234.5')
    await call(page, handle, 'setSelection', { top: 3, left: 0, bottom: 3, right: 0 })
    await page.locator(toolbarSel).getByRole('button', { name: '数字格式' }).click()
    const card = page.locator('.ooxml-numfmt-card')
    await expect(card).toBeVisible()
    const codeInput = card.locator('[data-code]')
    await codeInput.fill('#,##0.00')
    await expect(card.locator('[data-preview]')).toContainText('1,234.50') // 实时预览
    await card.locator('[data-ok]').click()
    await expect(card).toBeHidden()
    expect(await call(page, handle, 'getCellText', 3, 0)).toBe('1,234.50')
  })
}

test.describe('数字格式编辑 e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.action-toolbar', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-action-toolbar', '__excelViewerReact')
  run('Vue2', 'http://localhost:5302/', 'canvas.ov-grid-canvas', '.ov-action-toolbar', '__excelViewerVue2')
})
