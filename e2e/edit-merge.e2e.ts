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
  await page.waitForFunction((h) => (window as any)[h] != null, handle, { timeout: 5_000 })
  await page.getByText('编辑模式').click()
}

const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

// G1:合并单元格 —— 合并首数据行 (2,0)-(2,2):锚点保留、被覆盖格清空;undo 还原;拆分移除合并。
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: mergeCells 清被覆盖格 + undo 还原;unmergeCells 拆分`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    const anchor = await call(page, handle, 'getCellValue', 2, 0) // A3 笔记本电脑(锚点)
    const covered = await call(page, handle, 'getCellValue', 2, 1) // B3 单价(将被清)
    expect(covered).not.toBeNull()

    // 合并 (2,0)-(2,2)
    expect(await call(page, handle, 'mergeCells', { top: 2, left: 0, bottom: 2, right: 2 })).toBe(true)
    expect(await call(page, handle, 'getCellValue', 2, 0)).toBe(anchor) // 锚点保留
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBeNull() // 被覆盖格清空
    expect(await call(page, handle, 'getCellValue', 2, 2)).toBeNull()

    // undo → 被清格 + 合并都还原
    await call(page, handle, 'undo')
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBe(covered)

    // 再合并 → 拆分(拆分只移除合并,不恢复值);单格不合并返回 false
    await call(page, handle, 'mergeCells', { top: 2, left: 0, bottom: 2, right: 2 })
    expect(await call(page, handle, 'unmergeCells', { top: 2, left: 0, bottom: 2, right: 2 })).toBe(true)
    expect(await call(page, handle, 'mergeCells', { top: 5, left: 5, bottom: 5, right: 5 })).toBe(false)
  })
}

test.describe('合并单元格 e2e(G1:merge/unmerge + undo)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
