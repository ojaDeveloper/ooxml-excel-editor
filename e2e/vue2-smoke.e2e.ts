import { test, expect, type Page } from '@playwright/test'

/**
 * Vue 2 壳回归网(1.8.0 新增)—— 此前 Vue 2 完全没有 e2e 覆盖,改 Vue 2 全靠手测。
 * Vue 2 壳的高风险点都是"DOM 复用/imperative DOM"相关(见 src/vue2/ExcelViewer.ts 头部 doc):
 *  - patch 复用 controller 持有的 canvas/scroller DOM → controller stale(key + createElement 兜底)
 *  - Vue 2.6 函数 ref 被忽略 → renderArea 拿不到(string ref + Object.assign 暴露 API)
 * 这套 smoke 把这些"只有真 Vue 2 浏览器才暴露"的回归钉死。跑在独立 5302 dev server(plugin-vue2)。
 *
 * 跑法:npx playwright test vue2-smoke(playwright.config.ts 已配 5302 webServer)。
 */
const URL = 'http://localhost:5302/'
const CANVAS = 'canvas.ov-grid-canvas'
const HANDLE = '__excelViewerVue2'

const call = (page: Page, fn: string, ...args: unknown[]) =>
  page.evaluate(([f, a]) => (window as any).__excelViewerVue2[f as string](...(a as unknown[])), [fn, args] as const)

async function loadSample(page: Page) {
  await page.goto(URL)
  await page.getByRole('button', { name: '加载示例' }).click()
  await expect(page.locator(CANVAS)).toBeVisible({ timeout: 20_000 })
  await page.waitForFunction((h) => !!(window as any)[h]?.getWorkbook?.()?.sheets?.length, HANDLE, { timeout: 20_000 })
}

test.describe('Vue 2 壳 smoke e2e', () => {
  test('加载示例 → canvas 渲染 + 模型有 sheet(渲染链路通,DOM 未被 patch 复用毁掉)', async ({ page }) => {
    await loadSample(page)
    const sheets = await call(page, 'getWorkbook')
    expect((sheets as any).sheets.length).toBeGreaterThan(0)
    // canvas 真有像素尺寸(controller 拿到了 DOM 并渲染,而非空壳)
    const box = await page.locator(CANVAS).boundingBox()
    expect(box!.width).toBeGreaterThan(100)
    expect(box!.height).toBeGreaterThan(100)
  })

  test('编辑模式 → editCell 写入 + getCellValue 读回 + undo 回退(命令栈不因 stale 失效)', async ({ page }) => {
    await loadSample(page)
    await page.locator('.edit-toggle').click()
    await call(page, 'setSelection', { top: 0, left: 0, bottom: 0, right: 0 })
    const before = await call(page, 'getCellValue', 0, 0)
    await call(page, 'editCell', 0, 0, 'VUE2_E2E')
    expect(await call(page, 'getCellValue', 0, 0)).toBe('VUE2_E2E')
    await call(page, 'undo')
    expect(await call(page, 'getCellValue', 0, 0)).toBe(before)
  })

  test('rectOf 几何 + setSelection 不抛(选区/命中检测在 imperative DOM 上正常)', async ({ page }) => {
    await loadSample(page)
    await call(page, 'setSelection', { top: 2, left: 1, bottom: 2, right: 1 })
    const rect = (await call(page, 'rectOf', 2, 1)) as { x: number; y: number; w: number; h: number } | null
    expect(rect).toBeTruthy()
    expect(rect!.w).toBeGreaterThan(0)
    expect(rect!.h).toBeGreaterThan(0)
  })

  test('demo 顶栏按钮 1:1 渲染(加载示例 / JSON 示例 / 编辑模式 toggle 都在)', async ({ page }) => {
    await page.goto(URL)
    await expect(page.getByRole('button', { name: '加载示例' })).toBeVisible()
    await expect(page.getByRole('button', { name: 'JSON 示例' })).toBeVisible()
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(CANVAS)).toBeVisible({ timeout: 20_000 })
    await expect(page.locator('.edit-toggle')).toBeVisible() // 数据载入后编辑开关出现
  })
})
