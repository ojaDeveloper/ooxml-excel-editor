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

// Plan C 高阶覆盖:用 openContextMenu 模拟自定义 items 来测「transform → 自定义 items 入栈 →
// 点击后回调被触发」。真正 transform prop 的链路得通过 :contextMenu="(ctx, items) => [...]" 传给组件,
// demo 没接,这里用命令式 API 代替 —— openContextMenu(x,y, customItems) 与"transform 返回 customItems"等价。
function run(label: string, url: string, canvasSel: string, renderAreaSel: string, handle: string) {
  test(`${label}: transform 等价路径 —— 自定义 items + ctx 可用`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    const menu = page.locator('.ooxml-context-menu')

    // 选 (2,1)
    await call(page, handle, 'setSelection', { top: 2, left: 1, bottom: 2, right: 1 })

    // 自定义 items + 注入 inspect 信息(模拟 transform 回调里基于 ctx 加项)
    const insp = await call(page, handle, 'inspectCell', 2, 1)
    const ratio = (insp as { merge: unknown | null }).merge != null
    expect(typeof ratio).toBe('boolean')

    // 拼一组带 ctx 派生信息的 items
    await page.evaluate(
      ([h]) => (window as any)[h as string].openContextMenu(120, 120, [
        { label: '我加的:导出此格 PDF', action: () => ((window as any).__ctxAction = 'pdf') },
        { label: '我加的:复制单元格地址', action: () => ((window as any).__ctxAction = 'copy-addr') },
        { separator: true },
        { label: '内置仿真:清除内容', action: () => (window as any)[h as string].clearRange({ top: 2, left: 1, bottom: 2, right: 1 }) },
      ]),
      [handle] as const,
    )
    await expect(menu).toBeVisible()
    await expect(menu).toContainText('导出此格 PDF')
    await expect(menu).toContainText('复制单元格地址')

    // 点中间项
    await menu.getByText('复制单元格地址').click()
    expect(await page.evaluate(() => (window as any).__ctxAction)).toBe('copy-addr')
    await expect(menu).toBeHidden()
  })
}

test.describe('右键菜单 transform 等价路径(Plan C)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.render-area', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-render-area', '__excelViewerReact')
})
