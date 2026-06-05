import { test, expect, type Page } from '@playwright/test'

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
  await page.waitForFunction((h) => (window as any)[h] != null, handle, { timeout: 20_000 })
}

const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

// Phase A:向下滚动自动出现虚拟空行,但不写进 dimension/导出
function run(label: string, url: string, canvasSel: string, handle: string, scrollerSel: string, spacerSel: string) {
  test(`${label}: 滚动出虚拟空行 + spacer 增长 + 导出不含虚拟行`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    const dimRows = (await page.evaluate((h) => {
      const v = (window as any)[h]
      return v.getWorkbook().sheets[v.getActiveSheet()].dimension.rows
    }, handle)) as number

    const spacerH0 = await page.locator(spacerSel).evaluate((el) => (el as HTMLElement).offsetHeight)

    // 滚到底两次(grow-only:每次延伸出下一屏空行)
    for (let i = 0; i < 2; i++) {
      await page.locator(scrollerSel).evaluate((el) => {
        el.scrollTop = el.scrollHeight
      })
      await page.waitForTimeout(120)
    }

    const ext = (await call(page, handle, 'getVirtualExtent')) as { rows: number; cols: number }
    expect(ext.rows).toBeGreaterThan(dimRows) // 虚拟范围已超过数据行
    const spacerH1 = await page.locator(spacerSel).evaluate((el) => (el as HTMLElement).offsetHeight)
    expect(spacerH1).toBeGreaterThan(spacerH0) // spacer 撑大 → 能继续往下滚

    // 导出/数据安全:getSheetData 仍按 dimension,不含成千上万虚拟空行
    const data = (await call(page, handle, 'getSheetData')) as unknown[]
    expect(data.length).toBe(dimRows)

    // dimension 未被虚拟滚动改动(只有真编辑才增长)
    const dimRowsAfter = (await page.evaluate((h) => {
      const v = (window as any)[h]
      return v.getWorkbook().sheets[v.getActiveSheet()].dimension.rows
    }, handle)) as number
    expect(dimRowsAfter).toBe(dimRows)
  })
}

test.describe('虚拟空行(滚动自动延伸,不动 dimension)e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer', '.scroller', '.scroller > .spacer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact', '.rxl-scroller', '.rxl-scroller > .rxl-spacer')
})
