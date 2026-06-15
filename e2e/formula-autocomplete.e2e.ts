import { test, expect, type Page } from '@playwright/test'

/**
 * 公式自动补全 e2e(1.14.0)。三壳共测 —— 补全在框架无关默认编辑器里,三壳自动都有。
 * 编辑 B4(非 A 列,用默认编辑器)→ 输 `=SU` → 弹函数列表(含 SUM)→ 点 SUM → 变 `=SUM(`。
 */
const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

function run(label: string, url: string, canvasSel: string, renderAreaSel: string, handle: string) {
  test(`${label}: 公式补全 =SU → 弹 SUM → 点选插入 SUM(`, async ({ page }) => {
    await page.goto(url)
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => !!(window as any)[h]?.getWorkbook?.()?.sheets?.length, handle, { timeout: 20_000 })
    await page.locator('.edit-toggle').click()

    // 双击 B4(行3列1;非 A 列→默认编辑器有补全;非只读行)
    await call(page, handle, 'setSelection', { top: 3, left: 1, bottom: 3, right: 1 })
    const box = (await page.locator(renderAreaSel).boundingBox())!
    const r = (await call(page, handle, 'rectOf', 3, 1)) as { x: number; y: number; w: number; h: number }
    await page.mouse.dblclick(box.x + r.x + r.w / 2, box.y + r.y + r.h / 2)
    const editor = page.locator('textarea.ooxml-cell-editor')
    await expect(editor).toBeVisible()
    await editor.fill('=SU')

    const ac = page.locator('.ooxml-formula-ac')
    await expect(ac).toBeVisible()
    await expect(ac).toContainText('SUM')
    await ac.getByText('SUM', { exact: true }).first().click()
    await expect(editor).toHaveValue('=SUM(')
    await expect(ac).toBeHidden() // 接受后关闭
  })
}

test.describe('公式自动补全 e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.render-area', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-render-area', '__excelViewerReact')
  run('Vue2', 'http://localhost:5302/', 'canvas.ov-grid-canvas', '.ov-render-area', '__excelViewerVue2')
})
