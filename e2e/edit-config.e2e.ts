import { test, expect, type Page } from '@playwright/test'

/** 等画布真正绘制(不透明已绘制像素),并等命令式句柄就绪 */
async function ready(page: Page, canvasSel: string, handle: string) {
  await page.getByRole('button', { name: '加载示例' }).click()
  await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
  await page.waitForFunction(
    (sel) => {
      const c = document.querySelector(sel) as HTMLCanvasElement | null
      if (!c || !c.width) return false
      const d = c.getContext('2d')!.getImageData(0, 0, 200, 200).data
      for (let i = 0; i < d.length; i += 4) if (d[i + 3] === 255 && d[i] < 248) return true
      return false
    },
    canvasSel,
    { timeout: 20_000 },
  )
  await page.waitForFunction((h) => (window as any)[h] != null, handle, { timeout: 5_000 })
}

const editable = (page: Page, handle: string, r: number, c: number) =>
  page.evaluate(([h, rr, cc]) => (window as any)[h].isCellEditable(rr, cc), [handle, r, c] as const)

// 配置 + 只读判定(E0):默认只读;开编辑后数据格可编辑,只读区域(表头行)仍只读。
test.describe('编辑配置 e2e(E0:editable + readOnlyRanges)', () => {
  test('Vue: 默认只读 → 开编辑模式 → 数据格可编辑、表头行只读', async ({ page }) => {
    await page.goto('/')
    await ready(page, 'canvas.grid-canvas', '__excelViewer')

    // 默认:editable 未开 → 一律不可编辑
    expect(await editable(page, '__excelViewer', 2, 0)).toBe(false)

    // 勾选「编辑模式」→ 数据格(行2)可编辑,只读区域(表头行1)仍只读
    await page.getByText('编辑模式').click()
    await expect.poll(() => editable(page, '__excelViewer', 2, 0)).toBe(true)
    expect(await editable(page, '__excelViewer', 1, 0)).toBe(false) // readOnlyRanges 命中表头行
  })

  test('React: 默认只读 → 开编辑模式 → 数据格可编辑、表头行只读', async ({ page }) => {
    await page.goto('/react.html')
    await ready(page, 'canvas.rxl-canvas', '__excelViewerReact')

    expect(await editable(page, '__excelViewerReact', 2, 0)).toBe(false)

    await page.getByText('编辑模式').click()
    await expect.poll(() => editable(page, '__excelViewerReact', 2, 0)).toBe(true)
    expect(await editable(page, '__excelViewerReact', 1, 0)).toBe(false)
  })
})
