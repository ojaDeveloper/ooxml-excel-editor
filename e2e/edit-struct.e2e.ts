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
const win = (page: Page, key: string) => page.evaluate((k) => (window as any)[k], key)

// E7:行列结构编辑 —— 第 2 行(0-based)= 首数据行(B3=5999),第 3 行(B4=399)。
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: insertRows 下移 + struct-change + undo;deleteRows 上移 + undo`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    const b3 = (await call(page, handle, 'getCellValue', 2, 1)) as number // 5999
    const b4 = (await call(page, handle, 'getCellValue', 3, 1)) as number // 399
    expect(b3).not.toBe(b4)

    // 在第 2 行插入一行 → 第 2 行变空,原数据下移到第 3 行
    expect(await call(page, handle, 'insertRows', 2, 1)).toBe(true)
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBeNull() // 新空行
    expect(await call(page, handle, 'getCellValue', 3, 1)).toBe(b3) // 原 B3 下移
    expect(await win(page, '__lastStructChange')).toMatchObject({ op: 'insert-rows', at: 2, count: 1 })

    // undo → 还原
    await call(page, handle, 'undo')
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBe(b3)

    // 删除第 2 行 → 原第 3 行(B4)上移到第 2 行
    expect(await call(page, handle, 'deleteRows', 2, 1)).toBe(true)
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBe(b4)
    expect(await win(page, '__lastStructChange')).toMatchObject({ op: 'delete-rows' })

    // undo → 被删行还原
    await call(page, handle, 'undo')
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBe(b3)
    expect(await call(page, handle, 'getCellValue', 3, 1)).toBe(b4)
  })
}

test.describe('行列结构编辑 e2e(E7:增删行 + struct-change + undo)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
