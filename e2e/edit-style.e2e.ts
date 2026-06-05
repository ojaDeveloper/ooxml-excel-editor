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

// E5:样式编辑 —— setStyle 给单价格(2,1)加粗 → styleId 变 + cell-change 前后 style.font.bold 不同;undo 还原。
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: setStyle 加粗 → 前后 style 变 + cell-change;undo 还原`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    const cell = { top: 2, left: 1, bottom: 2, right: 1 } // 单价 B3
    const snap0 = (await call(page, handle, 'getCellSnapshot', 2, 1)) as any
    expect(snap0.style?.font?.bold).toBeFalsy() // 初始非粗

    expect(await call(page, handle, 'setStyle', cell, { font: { bold: true } })).toBe(true)
    const snap1 = (await call(page, handle, 'getCellSnapshot', 2, 1)) as any
    expect(snap1.style.font.bold).toBe(true) // 加粗生效
    // cell-change 前后 style 不同(底层结构事件)
    const evt = await win(page, '__lastCellChange')
    expect(evt.after.style.font.bold).toBe(true)
    expect(evt.before.style?.font?.bold).toBeFalsy()

    // undo → 还原非粗
    await call(page, handle, 'undo')
    const snap2 = (await call(page, handle, 'getCellSnapshot', 2, 1)) as any
    expect(snap2.style?.font?.bold).toBeFalsy()
  })
}

test.describe('样式编辑 e2e(E5:setStyle 加粗 + 前后快照 + undo)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
