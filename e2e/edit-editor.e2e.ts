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

const handleCall = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

// E2:editor 扩展钩子 —— 自定义 <select> 编辑器(框架无关 DOM),挂载/提交/取消/单活动/只读拦截。
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 自定义 select 编辑器 → 选值提交 + cell-change;Esc 取消;只读/无编辑器不弹`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // 第 0 列(产品)有自定义 select 编辑器 → beginEdit(2,0) 弹出
    expect(await handleCall(page, handle, 'beginEdit', 2, 0)).toBe(true)
    const sel = page.locator('select.demo-cell-editor')
    await expect(sel).toBeVisible()

    // 选 'AAA' → onchange → ctx.commit → 编辑器卸载 + 值变 + cell-change
    await sel.selectOption('AAA')
    await expect(sel).toBeHidden()
    expect(await handleCall(page, handle, 'getCellText', 2, 0)).toBe('AAA')
    const evt = await page.evaluate(() => (window as any).__lastCellChange)
    expect(evt.after.text).toBe('AAA')

    // Esc 取消:再编辑 → Esc → 编辑器关、值不变
    expect(await handleCall(page, handle, 'beginEdit', 2, 0)).toBe(true)
    await expect(sel).toBeVisible()
    await sel.press('Escape')
    await expect(sel).toBeHidden()
    expect(await handleCall(page, handle, 'getCellText', 2, 0)).toBe('AAA') // 仍是上次提交的

    // 只读格(表头行 1)→ beginEdit 返回 false、不弹
    expect(await handleCall(page, handle, 'beginEdit', 1, 0)).toBe(false)
    await expect(sel).toBeHidden()
    // 无自定义编辑器的列(col 2,demo 只给 col 0)→ E3 起用内置文本编辑器,不是自定义 select
    expect(await handleCall(page, handle, 'beginEdit', 2, 2)).toBe(true)
    await expect(sel).toBeHidden() // 自定义 select 不出现
    await expect(page.locator('input.ooxml-cell-editor')).toBeVisible() // 内置编辑器
    await handleCall(page, handle, 'cancelEdit')
  })
}

test.describe('编辑器扩展 e2e(E2:自定义 editor)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
