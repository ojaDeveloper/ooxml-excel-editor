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
  // 公式引擎异步 warm(动态 import hyperformula)→ 等就绪再编辑,保证级联确定性
  await page.waitForFunction((h) => (window as any)[h].isRecalcReady() === true, handle, { timeout: 15_000 })
}

const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)
const win = (page: Page, key: string) => page.evaluate((k) => (window as any)[k], key)

// E4:公式重算 —— 样例 D 列(0-based col 3)= B*C 公式。改数量(C,col 2)→ 金额(D)级联重算 + 发 cell-change;undo 反向重算。
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 改被引用格 → 公式格级联重算 + cell-change;undo 反向重算`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // 首数据行(0-based row 2):B3=单价, C3=数量, D3==B3*C3
    const b3 = (await call(page, handle, 'getCellValue', 2, 1)) as number // 单价 5999
    const c3 = (await call(page, handle, 'getCellValue', 2, 2)) as number // 数量 120
    expect(await call(page, handle, 'getCellValue', 2, 3)).toBe(b3 * c3) // D3 初值 = B*C

    // 改 C3=1000 → D3 级联到 B3*1000
    expect(await call(page, handle, 'editCell', 2, 2, 1000)).toBe(true)
    expect(await call(page, handle, 'getCellValue', 2, 3)).toBe(b3 * 1000) // 公式格重算
    // 最后一条 cell-change 应是依赖格 D3(col 3)的级联
    const evt = await win(page, '__lastCellChange')
    expect(evt.after.col).toBe(3)
    expect(evt.after.raw).toBe(b3 * 1000)
    expect(evt.before.raw).toBe(b3 * c3) // 前态 = 旧缓存值

    // undo → C3 回 120,D3 反向重算回 B3*120
    await call(page, handle, 'undo')
    expect(await call(page, handle, 'getCellValue', 2, 2)).toBe(c3)
    expect(await call(page, handle, 'getCellValue', 2, 3)).toBe(b3 * c3)
  })
}

test.describe('公式重算 e2e(E4:依赖格级联 + undo 反算)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
