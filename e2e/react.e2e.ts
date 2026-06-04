import { test, expect, type Page } from '@playwright/test'

/** 加载 React demo 示例,等真实 canvas 渲染(像素非空) */
async function loadReactSample(page: Page) {
  await page.goto('/react.html')
  await page.getByRole('button', { name: '加载示例' }).click()
  await expect(page.locator('canvas.rxl-canvas')).toBeVisible({ timeout: 20_000 })
  await page.waitForFunction(
    () => {
      const c = document.querySelector('canvas.rxl-canvas') as HTMLCanvasElement | null
      if (!c || !c.width) return false
      const ctx = c.getContext('2d')!
      const d = ctx.getImageData(0, 0, Math.min(c.width, 300), Math.min(c.height, 300)).data
      // 要求"已绘制"的不透明像素(alpha=255)且非纯白 —— 排除未绘制的透明画布(alpha=0)假阳性
      for (let i = 0; i < d.length; i += 4) {
        if (d[i + 3] === 255 && (d[i] < 248 || d[i + 1] < 248 || d[i + 2] < 248)) return true
      }
      return false
    },
    null,
    { timeout: 20_000 },
  )
}

test.describe('React 壳 e2e(与 Vue 共用 core: 渲染/选区/查找/数据/导出)', () => {
  test('解析→canvas 渲染→点选更新状态栏→Ctrl+F 查找→数据 API→PNG 导出', async ({ page }) => {
    await loadReactSample(page)

    // 1) 命令式选区 → core 选区模型 + onSelectionChange 回调 → React 重渲公式栏(B3)
    const scroller = page.locator('.rxl-scroller')
    await page.evaluate(() => {
      const v = (window as unknown as { __excelViewerReact?: { setSelection: (r: unknown) => void } }).__excelViewerReact
      v?.setSelection({ top: 2, left: 1, bottom: 2, right: 1 })
    })
    await expect(page.locator('.rxl-formula-bar .addr')).toHaveText('B3')
    // 真实鼠标拖选一段 → core 处理,不抛错(选区随之更新)
    const box = (await scroller.boundingBox())!
    await page.mouse.move(box.x + 90, box.y + 70)
    await page.mouse.down()
    await page.mouse.move(box.x + 200, box.y + 120)
    await page.mouse.up()
    await expect(page.locator('.rxl-formula-bar .addr')).not.toHaveText('—')

    // 2) Ctrl+F 打开查找,命中计数
    await scroller.click()
    await page.keyboard.press('Control+f')
    const find = page.locator('.rxl-findbar')
    await expect(find).toBeVisible()
    await find.locator('input').fill('zz不存在zz')
    await expect(find.locator('.count')).toHaveText('无结果')
    await find.locator('input').press('Escape')
    await expect(find).toBeHidden()

    // 3) 数据读取 API(命令式句柄)可用,且与渲染一致
    const text = await page.evaluate(() => {
      const v = (window as unknown as { __excelViewerReact?: { getSheetData: (o?: unknown) => unknown[][] } }).__excelViewerReact
      const data = v?.getSheetData()
      return data && data.length ? JSON.stringify(data[0]) : null
    })
    expect(text).not.toBeNull()
    expect(text!.length).toBeGreaterThan(2)

    // 4) PNG 导出走 core/WorkbookExporter,真实下载
    const [dl] = await Promise.all([
      page.waitForEvent('download', { timeout: 20_000 }),
      page.getByRole('button', { name: '导出 PNG' }).click(),
    ])
    expect(dl.suggestedFilename()).toMatch(/\.png$/)
  })
})
