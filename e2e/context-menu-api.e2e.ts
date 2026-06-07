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

async function cellCenter(page: Page, renderAreaSel: string, handle: string, row: number, col: number) {
  const box = (await page.locator(renderAreaSel).boundingBox())!
  const r = (await call(page, handle, 'rectOf', row, col)) as { x: number; y: number; w: number; h: number }
  return { x: box.x + r.x + r.w / 2, y: box.y + r.y + r.h / 2 }
}

// Plan C:右键菜单 API:
//  ① openContextMenu(x, y, items?) 命令式打开
//  ② closeContextMenu() 命令式关闭
//  ③ 内置 items 顺序确认(供 transform 拿到)
function run(label: string, url: string, canvasSel: string, renderAreaSel: string, handle: string) {
  test(`${label}: openContextMenu(自定义 items)→ 菜单出现;closeContextMenu 关掉`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    const menu = page.locator('.ooxml-context-menu')

    // 先选个格,让 ctx 有效
    const c = await cellCenter(page, renderAreaSel, handle, 2, 1)
    // 命令式打开,传自定义 items
    await page.evaluate(
      ([h, x, y]) => (window as any)[h as string].openContextMenu(x as number, y as number, [
        { label: '我是自定义项 A', action: () => ((window as any).__ctxFired = 'A') },
        { separator: true },
        { label: '我是自定义项 B', action: () => ((window as any).__ctxFired = 'B') },
      ]),
      [handle, c.x, c.y] as const,
    )
    await expect(menu).toBeVisible()
    await expect(menu).toContainText('我是自定义项 A')
    await expect(menu).toContainText('我是自定义项 B')
    // 点 A → 触发 action 并关闭菜单
    await menu.getByText('我是自定义项 A').click()
    await expect(menu).toBeHidden()
    expect(await page.evaluate(() => (window as any).__ctxFired)).toBe('A')

    // 再开一次,然后命令式关闭
    await page.evaluate(
      ([h, x, y]) => (window as any)[h as string].openContextMenu(x as number, y as number, [
        { label: '将被命令式关掉' },
      ]),
      [handle, c.x, c.y] as const,
    )
    await expect(menu).toBeVisible()
    await call(page, handle, 'closeContextMenu')
    await expect(menu).toBeHidden()
  })

  test(`${label}: openContextMenu 不传 items → 走当前选区算的内置 items`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    const menu = page.locator('.ooxml-context-menu')

    // 选 B3
    await call(page, handle, 'setSelection', { top: 2, left: 1, bottom: 2, right: 1 })
    // openContextMenu 不传 items → 内置(含"合并单元格"、"清除内容"等)
    await page.evaluate(([h]) => (window as any)[h as string].openContextMenu(100, 100), [handle] as const)
    await expect(menu).toBeVisible()
    await expect(menu).toContainText('清除内容')
    await call(page, handle, 'closeContextMenu')
  })
}

test.describe('右键菜单 API(Plan C:openContextMenu / closeContextMenu)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.render-area', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-render-area', '__excelViewerReact')
})
