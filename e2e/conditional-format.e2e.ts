import { test, expect, type Page } from '@playwright/test'

/**
 * 条件格式可编辑 e2e(1.9.0)。三壳共测:
 *  - 编程 API:addConditionalRule → 规则数 +1 → undo 回退(命令栈 + live render 不崩)
 *  - 对话框:openConditionalFormatDialog → 新建 cellIs 规则 → 保存 → 完成 → 规则数 +1
 * 条件格式对话框是框架无关 DOM(三壳共用一份),三壳跑同一套交互即验证 1:1。
 */
const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)
const cfLen = (page: Page, handle: string) =>
  page.evaluate((h) => (window as any)[h].getConditionalRules().length, handle)

async function loadAndEdit(page: Page, url: string, canvasSel: string, handle: string) {
  await page.goto(url)
  await page.getByRole('button', { name: '加载示例' }).click()
  await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
  await page.waitForFunction((h) => !!(window as any)[h]?.getWorkbook?.()?.sheets?.length, handle, { timeout: 20_000 })
  await page.locator('.edit-toggle').click() // 编辑模式(条件格式编辑需 editable)
}

function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 条件格式 API 新增规则 → 渲染不崩 → undo 回退`, async ({ page }) => {
    await loadAndEdit(page, url, canvasSel, handle)
    const before = await cfLen(page, handle)
    const id = await call(page, handle, 'addConditionalRule', { ranges: [{ top: 0, left: 0, bottom: 4, right: 0 }], type: 'cellIs', operator: 'greaterThan', formulae: ['50'], style: { fill: { type: 'solid', fgColor: '#FFEB9C' }, font: { color: '#9C5700' } } })
    expect(id).toBeTruthy()
    expect(await cfLen(page, handle)).toBe(before + 1)
    await expect(page.locator(canvasSel)).toBeVisible() // 加规则后仍正常渲染
    await call(page, handle, 'undo')
    expect(await cfLen(page, handle)).toBe(before) // 可撤销
  })

  test(`${label}: 条件格式对话框 新建 cellIs 规则 → 保存 → 完成`, async ({ page }) => {
    await loadAndEdit(page, url, canvasSel, handle)
    const before = await cfLen(page, handle)
    await call(page, handle, 'setSelection', { top: 0, left: 0, bottom: 4, right: 0 })
    await call(page, handle, 'openConditionalFormatDialog')
    const card = page.locator('.ooxml-cf-card')
    await expect(card).toBeVisible()
    await card.locator('[data-add]').click() // 新建规则
    await expect(card.locator('[data-type]')).toBeVisible() // 进编辑器
    await card.locator('[data-f="v0"]').fill('80') // cellIs 大于 80
    await card.locator('[data-save]').click() // 保存 → 回列表
    await card.locator('[data-done]').click() // 完成 → 应用
    await expect(card).toBeHidden()
    expect(await cfLen(page, handle)).toBe(before + 1)
  })
}

test.describe('条件格式可编辑 e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
  run('Vue2', 'http://localhost:5302/', 'canvas.ov-grid-canvas', '__excelViewerVue2')
})
