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

// G3:右键上下文菜单 —— 右键格 → 菜单出现(插入/删除/合并/清除)→ 点「清除内容」生效 + 菜单关闭。
function run(label: string, url: string, canvasSel: string, renderAreaSel: string, handle: string) {
  test(`${label}: 右键弹菜单 → 清除内容生效 + 菜单关闭;点外部关闭`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    const menu = page.locator('.ooxml-context-menu')

    expect(await call(page, handle, 'getCellValue', 2, 1)).not.toBeNull() // B3 有值

    // 右键单价格(2,1)
    const c = await cellCenter(page, renderAreaSel, handle, 2, 1)
    await page.mouse.click(c.x, c.y, { button: 'right' })
    await expect(menu).toBeVisible()
    await expect(menu).toContainText('合并单元格')
    await expect(menu).toContainText('删除 1 行')

    // 点「清除内容」→ 该格清空 + 菜单关闭
    await menu.getByText('清除内容').click()
    await expect(menu).toBeHidden()
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBeNull()

    // 再右键 → 点菜单外部关闭
    await page.mouse.click(c.x, c.y, { button: 'right' })
    await expect(menu).toBeVisible()
    const far = await cellCenter(page, renderAreaSel, handle, 6, 3) // 菜单外的另一格
    await page.mouse.click(far.x, far.y)
    await expect(menu).toBeHidden()
  })
}

test.describe('右键菜单 e2e(G3:插入/删除/合并/清除)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.render-area', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-render-area', '__excelViewerReact')
})
