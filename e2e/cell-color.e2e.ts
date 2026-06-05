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
  await page.getByText('编辑模式').click()
}

const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

// 背景色 / 字体色:回显当前格 + 改选区 + undo
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 背景色/字体色 回显 + 修改 + undo`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // 选可编辑格 A6
    await call(page, handle, 'setSelection', { top: 5, left: 0, bottom: 5, right: 0 })

    // 设背景红 → 回显 #FF0000
    expect(await call(page, handle, 'setSelectionFill', '#FF0000')).toBe(true)
    expect(await call(page, handle, 'getActiveFillColor')).toBe('#FF0000')

    // 设字体蓝 → 回显 #0000FF
    expect(await call(page, handle, 'setSelectionFontColor', '#0000FF')).toBe(true)
    expect(await call(page, handle, 'getActiveFontColor')).toBe('#0000FF')

    // 清除填充 → 回显默认白
    expect(await call(page, handle, 'setSelectionFill', null)).toBe(true)
    expect(await call(page, handle, 'getActiveFillColor')).toBe('#FFFFFF')

    // undo → 背景回到红(清除被撤销)
    await call(page, handle, 'undo')
    expect(await call(page, handle, 'getActiveFillColor')).toBe('#FF0000')
  })
}

test.describe('背景色 / 字体色 回显 + 修改 e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
