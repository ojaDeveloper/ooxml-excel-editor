import { test, expect, type Page } from '@playwright/test'

/**
 * 不连续多区域选择 e2e(1.13.0)。三壳共测 —— 纯 core 交互(Ctrl+点击行头/格 加选),壳只转发鼠标。
 * 点行3表头 → Ctrl+点行5表头 → 两个不相邻行区都选中(getSelectionRanges=2);普通点击回单选。
 */
const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

function run(label: string, url: string, canvasSel: string, renderAreaSel: string, handle: string) {
  test(`${label}: Ctrl+点击行头 多选两不邻行 → 2 区域;普通点击回单选`, async ({ page }) => {
    await page.goto(url)
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => !!(window as any)[h]?.getWorkbook?.()?.sheets?.length, handle, { timeout: 20_000 })

    const box = (await page.locator(renderAreaSel).boundingBox())!
    const r2 = (await call(page, handle, 'rectOf', 2, 0)) as { x: number; y: number; w: number; h: number }
    const r4 = (await call(page, handle, 'rectOf', 4, 0)) as { x: number; y: number; w: number; h: number }
    const HX = box.x + 12 // 行头区(x < rowHeaderWidth)

    // 点行3表头(选整行)
    await page.mouse.click(HX, box.y + r2.y + r2.h / 2)
    expect((await call(page, handle, 'getSelectionRanges') as unknown[]).length).toBe(1)

    // Ctrl+点行5表头 → 加选(不相邻)
    await page.keyboard.down('Control')
    await page.mouse.click(HX, box.y + r4.y + r4.h / 2)
    await page.keyboard.up('Control')
    const ranges = (await call(page, handle, 'getSelectionRanges')) as Array<{ top: number; bottom: number }>
    expect(ranges.length).toBe(2)
    expect(await call(page, handle, 'hasMultiSelection')).toBe(true)
    expect(ranges.map((r) => r.top).sort()).toEqual([2, 4]) // 第3、第5 行(0-based 2、4)

    // 普通点击行3 → 回单选
    await page.mouse.click(HX, box.y + r2.y + r2.h / 2)
    expect((await call(page, handle, 'getSelectionRanges') as unknown[]).length).toBe(1)
    expect(await call(page, handle, 'hasMultiSelection')).toBe(false)
  })
}

test.describe('不连续多区域选择 e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.render-area', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-render-area', '__excelViewerReact')
  run('Vue2', 'http://localhost:5302/', 'canvas.ov-grid-canvas', '.ov-render-area', '__excelViewerVue2')
})
