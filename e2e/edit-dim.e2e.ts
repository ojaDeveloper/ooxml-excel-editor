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

// E3.5:resize 入命令栈(可撤销 + dim-change + 脏状态)+ resetToOriginal 还原值与列宽。
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: setColumnWidth → dim-change + isDirty + undo 还原;resetToOriginal 还原值与列宽`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    const before = ((await call(page, handle, 'rectOf', 2, 2)) as { w: number }).w
    expect(await call(page, handle, 'isDirty')).toBe(false)

    // 程序化改列宽 → 列变宽 + isDirty + dim-change(after 精确=模型值)+ dirty-change
    const target = Math.round(before + 100)
    // Phase B 2026-06-08: setColumnWidth 返回成功条数 (number), 而非 boolean. 单值返 1.
    expect(await call(page, handle, 'setColumnWidth', 2, target)).toBe(1)
    const after = ((await call(page, handle, 'rectOf', 2, 2)) as { w: number }).w
    expect(after).toBeGreaterThan(before + 50) // 明显变宽(demo zoom=1)
    expect(await call(page, handle, 'isDirty')).toBe(true)
    expect((await win(page, '__lastDimChange')).after).toBe(target)
    expect((await win(page, '__lastDimChange')).axis).toBe('col')
    expect((await win(page, '__lastDirtyChange')).dirty).toBe(true)

    // 撤销 resize → 列宽还原(Ctrl+Z 走同一命令栈)
    await call(page, handle, 'undo')
    const undone = ((await call(page, handle, 'rectOf', 2, 2)) as { w: number }).w
    expect(Math.abs(undone - before)).toBeLessThanOrEqual(2)

    // 改一格值 + 再改列宽 → resetToOriginal 一键还原两者 + 清脏
    expect(await call(page, handle, 'editCell', 3, 3, 'RESET-TEST')).toBe(true)
    expect(await call(page, handle, 'getCellText', 3, 3)).toBe('RESET-TEST')
    await call(page, handle, 'setColumnWidth', 2, target)
    expect(await call(page, handle, 'isDirty')).toBe(true)

    expect(await call(page, handle, 'resetToOriginal')).toBe(true)
    expect(await call(page, handle, 'getCellText', 3, 3)).not.toBe('RESET-TEST') // 值还原
    const reset = ((await call(page, handle, 'rectOf', 2, 2)) as { w: number }).w
    expect(Math.abs(reset - before)).toBeLessThanOrEqual(2) // 列宽还原
    expect(await call(page, handle, 'isDirty')).toBe(false)
    expect((await win(page, '__lastDirtyChange')).dirty).toBe(false)
  })
}

test.describe('维度编辑 + 脏状态 e2e(E3.5:resize 入命令栈 + 还原原件)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
