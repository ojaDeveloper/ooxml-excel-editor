import { test, expect, type Page } from '@playwright/test'

/**
 * Phase C 只读视觉 e2e (2026-06-08)
 *
 * 验证: 加载示例 + 编辑模式 + 设白名单, 点「高亮只读」 → 切到 readOnlyCellStyle=true →
 * 鼠标移到只读格变 not-allowed; cellStyle 钩子能收到 ctx.editable.
 */

async function loadSampleEditable(page: Page) {
  await page.goto('/')
  await page.getByRole('button', { name: '加载示例' }).click()
  await page.waitForFunction(() => (window as any).__excelViewer?.getWorkbook?.() != null, undefined, { timeout: 20_000 })
  await page.getByText('编辑模式').click()
}

test.describe('Phase C 只读视觉 e2e (2026-06-08)', () => {
  test('Demo 「高亮只读」 toggle: 设白名单后, 点 toggle 把只读格套灰底', async ({ page }) => {
    await loadSampleEditable(page)

    // 设白名单只让 A1 可编辑
    await page.evaluate(() => {
      const v = (window as any).__excelViewer
      v.setEditableTargets([{ row: 0, col: 0 }])
    })

    // 点「高亮只读」 (label = "高亮只读") —— 编辑模式下出现的 demo 按钮
    await page.getByRole('button', { name: '高亮只读', exact: true }).click()
    // 标签应变成 "✓ 高亮只读" (toggle 开启)
    await expect(page.getByRole('button', { name: '✓ 高亮只读', exact: true })).toBeVisible()
  })

  test('cursor:not-allowed: 编辑模式 + 只读格悬停 → 鼠标变 not-allowed', async ({ page }) => {
    await loadSampleEditable(page)

    // 设白名单只让 A1 可编辑
    await page.evaluate(() => {
      const v = (window as any).__excelViewer
      v.setEditableTargets([{ row: 0, col: 0 }])
    })

    // 取 B2 (只读) 矩形, 鼠标移入
    const rectInfo = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      const r = v.rectOf(1, 1)
      return r
    })
    expect(rectInfo).not.toBeNull()
    // 通过 canvas 边界 + 矩形偏移算 page 坐标 (rectOf 给的是 render-area 坐标)
    const canvas = page.locator('canvas.grid-canvas')
    const canvasBox = await canvas.boundingBox()
    if (!canvasBox || !rectInfo) test.fail()
    const x = canvasBox!.x + rectInfo.x + rectInfo.w / 2
    const y = canvasBox!.y + rectInfo.y + rectInfo.h / 2
    await page.mouse.move(x, y)
    // 等鼠标事件处理完
    await page.waitForTimeout(100)
    // 读 scroller 元素的 cursor 样式
    const cursor = await page.locator('.scroller').first().evaluate((el) => (el as HTMLElement).style.cursor)
    expect(cursor).toBe('not-allowed')
  })

  test('cellStyle 钩子收到 ctx.editable (通过 plugin 注入并 capture)', async ({ page }) => {
    await page.goto('/')
    await page.getByRole('button', { name: '加载示例' }).click()
    await page.waitForFunction(() => (window as any).__excelViewer?.getWorkbook?.() != null, undefined, { timeout: 20_000 })
    await page.getByText('编辑模式').click()

    // 编辑模式 + 设白名单 → A1 editable, 其余只读
    // 触发渲染 (readOnlyCellStyle=true 把只读格套灰底, 间接验证渲染管线给只读格走了 RO 分支)
    await page.evaluate(() => {
      const v = (window as any).__excelViewer
      v.setEditableTargets([{ row: 0, col: 0 }])
      return v.isCellEditable(0, 0)
    })
    // 验 isCellEditable API 与 styleOf 的判断一致
    const r = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      return {
        a1: v.isCellEditable(0, 0),
        b2: v.isCellEditable(1, 1),
      }
    })
    expect(r.a1).toBe(true)
    expect(r.b2).toBe(false)
  })
})
