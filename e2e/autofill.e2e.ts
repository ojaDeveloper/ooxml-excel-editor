import { test, expect, type Page } from '@playwright/test'

/**
 * 自动填充柄 e2e(1.10.0)。三壳共测 —— 填充是纯 core canvas 交互(渲染层画柄 + 控制器拖拽),
 * 壳只转发鼠标事件,故三壳跑同一套即验证一致。
 * 用编辑模式把 A1=1、A2=2 后,拖 A1:A2 选区右下角填充柄到 A5 → 期望 A3/A4/A5 = 3/4/5;undo 回退。
 */
const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

function run(label: string, url: string, canvasSel: string, renderAreaSel: string, handle: string) {
  test(`${label}: 拖填充柄 等差序列(A3=1,A4=2 → A5..A7=3,4,5)+ undo`, async ({ page }) => {
    await page.goto(url)
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => !!(window as any)[h]?.getWorkbook?.()?.sheets?.length, handle, { timeout: 20_000 })
    await page.locator('.edit-toggle').click() // 编辑模式(填充柄需 editable)

    // 写入种子 1、2 到 A3/A4(避开 demo 的只读行 A2:E2),选中 A3:A4
    await call(page, handle, 'editCell', 2, 0, '1')
    await call(page, handle, 'editCell', 3, 0, '2')
    await call(page, handle, 'setSelection', { top: 2, left: 0, bottom: 3, right: 0 })

    // 填充柄在 A4(行3列0)右下角;拖到 A7(行6列0)中心
    const box = (await page.locator(renderAreaSel).boundingBox())!
    const a4 = (await call(page, handle, 'rectOf', 3, 0)) as { x: number; y: number; w: number; h: number }
    const a7 = (await call(page, handle, 'rectOf', 6, 0)) as { x: number; y: number; w: number; h: number }
    await page.mouse.move(box.x + a4.x + a4.w - 1, box.y + a4.y + a4.h - 1) // 落在填充柄
    await page.mouse.down()
    await page.mouse.move(box.x + a7.x + a7.w / 2, box.y + a7.y + a7.h / 2, { steps: 6 }) // 拖到 A7
    await page.mouse.up()

    expect(await call(page, handle, 'getCellValue', 4, 0)).toBe(3)
    expect(await call(page, handle, 'getCellValue', 5, 0)).toBe(4)
    expect(await call(page, handle, 'getCellValue', 6, 0)).toBe(5)

    await call(page, handle, 'undo') // 填充整体单次撤销
    expect(await call(page, handle, 'getCellValue', 4, 0)).not.toBe(3)
  })

  test(`${label}: Ctrl 拖单个数字 → 递增序列(A3=7 + Ctrl → A4/A5=8/9)`, async ({ page }) => {
    await page.goto(url)
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => !!(window as any)[h]?.getWorkbook?.()?.sheets?.length, handle, { timeout: 20_000 })
    await page.locator('.edit-toggle').click()

    await call(page, handle, 'editCell', 2, 0, '7') // A3 = 7(避开只读行 A2:E2)
    await call(page, handle, 'setSelection', { top: 2, left: 0, bottom: 2, right: 0 })
    const box = (await page.locator(renderAreaSel).boundingBox())!
    const a3 = (await call(page, handle, 'rectOf', 2, 0)) as { x: number; y: number; w: number; h: number }
    const a5 = (await call(page, handle, 'rectOf', 4, 0)) as { x: number; y: number; w: number; h: number }
    await page.mouse.move(box.x + a3.x + a3.w - 1, box.y + a3.y + a3.h - 1)
    await page.keyboard.down('Control')
    await page.mouse.down()
    await page.mouse.move(box.x + a5.x + a5.w / 2, box.y + a5.y + a5.h / 2, { steps: 6 })
    await page.mouse.up()
    await page.keyboard.up('Control')

    expect(await call(page, handle, 'getCellValue', 3, 0)).toBe(8) // Ctrl → 递增而非复制
    expect(await call(page, handle, 'getCellValue', 4, 0)).toBe(9)
  })
}

test.describe('自动填充柄 e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.render-area', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-render-area', '__excelViewerReact')
  run('Vue2', 'http://localhost:5302/', 'canvas.ov-grid-canvas', '.ov-render-area', '__excelViewerVue2')
})
