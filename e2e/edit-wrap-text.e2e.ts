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

// 自动换行(WPS 风格 toggle):
//   ① 选区里 wrap 状态:'all' 全开 / 'none' 全关 / 'mixed' 混合
//   ② 全关 → 切换 → 全开;再切 → 全关;undo 还原
//   ③ 混合(部分开)→ 切换 → 全开(WPS 行为:有任一未开就全开,不切到全关)
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: toggleWrapTextOnSelection toggle + mixed → all + undo 还原`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // 单格 (2,1):初始无 wrap
    expect(await call(page, handle, 'setSelection', { top: 2, left: 1, bottom: 2, right: 1 })).toBeUndefined()
    expect(await call(page, handle, 'getSelectionWrapState')).toBe('none')

    // toggle ON
    expect(await call(page, handle, 'toggleWrapTextOnSelection')).toBe(true)
    expect(await call(page, handle, 'getSelectionWrapState')).toBe('all')
    const snap1 = (await call(page, handle, 'getCellSnapshot', 2, 1)) as any
    expect(snap1.style?.wrapText).toBe(true)

    // toggle OFF
    expect(await call(page, handle, 'toggleWrapTextOnSelection')).toBe(true)
    expect(await call(page, handle, 'getSelectionWrapState')).toBe('none')
    const snap2 = (await call(page, handle, 'getCellSnapshot', 2, 1)) as any
    expect(snap2.style?.wrapText).toBeFalsy()

    // undo → 回到 toggle ON 之后(wrap=true)
    await call(page, handle, 'undo')
    expect(await call(page, handle, 'getSelectionWrapState')).toBe('all')

    // 扩大选区 (2,1)→(3,2):一格 wrap=true(刚才那个),其余无 → mixed
    expect(await call(page, handle, 'setSelection', { top: 2, left: 1, bottom: 3, right: 2 })).toBeUndefined()
    expect(await call(page, handle, 'getSelectionWrapState')).toBe('mixed')

    // mixed → toggle:应"全开"(WPS 行为)
    expect(await call(page, handle, 'toggleWrapTextOnSelection')).toBe(true)
    expect(await call(page, handle, 'getSelectionWrapState')).toBe('all')
  })

  test(`${label}: 只读(editable=false)时 toggle 不生效`, async ({ page }) => {
    await page.goto(url)
    // 不进编辑模式
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => (window as any)[h] != null, handle, { timeout: 20_000 })
    await call(page, handle, 'setSelection', { top: 2, left: 1, bottom: 2, right: 1 })
    expect(await call(page, handle, 'toggleWrapTextOnSelection')).toBe(false)
    const snap = (await call(page, handle, 'getCellSnapshot', 2, 1)) as any
    expect(snap?.style?.wrapText).toBeFalsy()
  })
}

test.describe('自动换行 toggle e2e(WPS 风格 wrapText)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
