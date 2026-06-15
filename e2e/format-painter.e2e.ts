import { test, expect, type Page } from '@playwright/test'

/**
 * 格式刷 e2e(1.12.0)。三壳共测 —— 格式刷是纯 core 交互(控制器采样 + onMouseUp 刷),壳只转发事件 + 工具栏按钮。
 * 给源格 A4 设红底 → 选中 A4 点「格式刷」采样 → 点目标 A6 → A6 也变红底;Esc/再点退出。
 */
const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

function run(label: string, url: string, canvasSel: string, toolbarSel: string, renderAreaSel: string, handle: string) {
  test(`${label}: 格式刷 采样红底 → 刷到目标格`, async ({ page }) => {
    await page.setViewportSize({ width: 1680, height: 900 }) // 工具栏长,宽屏避免「格式刷」溢出进「更多」
    await page.goto(url)
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => !!(window as any)[h]?.getWorkbook?.()?.sheets?.length, handle, { timeout: 20_000 })
    await page.locator('.edit-toggle').click()

    // 源格 A4(3,0)设红底
    await call(page, handle, 'setSelection', { top: 3, left: 0, bottom: 3, right: 0 })
    await call(page, handle, 'setSelectionFill', '#FF0000')
    expect((await call(page, handle, 'getActiveFillColor') as string).toUpperCase()).toBe('#FF0000')

    // 点「格式刷」采样(源仍选中)
    await page.locator(toolbarSel).getByRole('button', { name: '格式刷' }).click()
    expect(await call(page, handle, 'isFormatPainterArmed')).toBe(true)

    // 点目标 A6(5,0)→ 刷上
    const box = (await page.locator(renderAreaSel).boundingBox())!
    const a6 = (await call(page, handle, 'rectOf', 5, 0)) as { x: number; y: number; w: number; h: number }
    await page.mouse.click(box.x + a6.x + a6.w / 2, box.y + a6.y + a6.h / 2)

    expect((await call(page, handle, 'getActiveFillColor') as string).toUpperCase()).toBe('#FF0000') // A6 已被刷红
    expect(await call(page, handle, 'isFormatPainterArmed')).toBe(false) // 单次刷完退出

    await call(page, handle, 'undo') // 刷上的样式单次撤销
    expect((await call(page, handle, 'getActiveFillColor') as string).toUpperCase()).not.toBe('#FF0000')
  })
}

test.describe('格式刷 e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.action-toolbar', '.render-area', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-action-toolbar', '.rxl-render-area', '__excelViewerReact')
  run('Vue2', 'http://localhost:5302/', 'canvas.ov-grid-canvas', '.ov-action-toolbar', '.ov-render-area', '__excelViewerVue2')
})
