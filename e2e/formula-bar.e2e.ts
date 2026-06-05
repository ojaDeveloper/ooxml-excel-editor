import { test, expect, type Page } from '@playwright/test'

async function ready(page: Page, canvasSel: string, handle: string) {
  await page.getByRole('button', { name: '加载示例' }).click()
  await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
  // 等渲染器真正绑定并画出内容(否则 setSelection 时 renderer 还未就绪)
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

// 公式栏(Fx 内容条)可编辑 + 与单元格联动
function run(label: string, url: string, canvasSel: string, handle: string, barSel: string) {
  test(`${label}: 公式栏输入 → 改活动格 + 回车下移;选区变 → 栏联动`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    const input = page.locator(`${barSel} .content-input`)

    // 选中可编辑格 A6 (row5,col0),公式栏出现可编辑 input
    await call(page, handle, 'setSelection', { top: 5, left: 0, bottom: 5, right: 0 })
    await expect(input).toBeVisible()

    // 在公式栏输入 42 → 回车 → A6 = 42(数字)+ 活动格下移到 A7
    await input.fill('42')
    await input.press('Enter')
    expect(await call(page, handle, 'getCellValue', 5, 0)).toBe(42)
    const active = (await call(page, handle, 'getSelection')) as { top: number; left: number }
    expect(active.top).toBe(6) // 回车后下移一行

    // 联动反向:选区切到 A6 → 公式栏 input 反映该格内容 "42"
    await call(page, handle, 'setSelection', { top: 5, left: 0, bottom: 5, right: 0 })
    await expect(input).toHaveValue('42')

    // 命令式改格(模拟格内编辑提交)→ 公式栏联动更新
    await call(page, handle, 'editCell', 5, 0, 'hello')
    await expect(input).toHaveValue('hello')
  })
}

test.describe('公式栏(Fx 内容条)可编辑 + 联动 e2e', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer', '.formula-bar')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact', '.rxl-formula-bar')
})
