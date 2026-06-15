import { test, expect, type Page } from '@playwright/test'

/**
 * 查找替换 e2e(1.11.0)。三壳共测 —— 查找栏开 editable 才显示替换行。
 * 编辑模式写两格 ZZZ → 打开查找栏 → 查 ZZZ、替换为 YYY → 全部替换 → 两格变 YYY;undo 回退。
 */
const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

function run(label: string, url: string, canvasSel: string, toolbarSel: string, findbarSel: string, handle: string) {
  test(`${label}: 查找栏 全部替换 ZZZ→YYY(两格)+ undo`, async ({ page }) => {
    await page.goto(url)
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => !!(window as any)[h]?.getWorkbook?.()?.sheets?.length, handle, { timeout: 20_000 })
    await page.locator('.edit-toggle').click() // 编辑模式(替换需 editable)

    // 写种子(避开只读行 A2:E2)
    await call(page, handle, 'editCell', 3, 0, 'ZZZ')
    await call(page, handle, 'editCell', 4, 0, 'ZZZ')

    // 打开查找栏
    await page.locator(toolbarSel).getByRole('button', { name: /查找/ }).click()
    const bar = page.locator(findbarSel)
    await expect(bar).toBeVisible()
    await bar.getByPlaceholder('查找…').fill('ZZZ')
    await bar.getByPlaceholder('替换为…').fill('YYY')
    await bar.getByRole('button', { name: '全部替换' }).click()

    expect(await call(page, handle, 'getCellValue', 3, 0)).toBe('YYY')
    expect(await call(page, handle, 'getCellValue', 4, 0)).toBe('YYY')

    await call(page, handle, 'undo') // 全部替换整体单次撤销
    expect(await call(page, handle, 'getCellValue', 3, 0)).toBe('ZZZ')
    expect(await call(page, handle, 'getCellValue', 4, 0)).toBe('ZZZ')
  })
}

test.describe('查找替换 e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.action-toolbar', '.find-bar', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-action-toolbar', '.rxl-findbar', '__excelViewerReact')
  run('Vue2', 'http://localhost:5302/', 'canvas.ov-grid-canvas', '.ov-action-toolbar', '.ov-findbar', '__excelViewerVue2')
})
