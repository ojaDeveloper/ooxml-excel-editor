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

// G2:粘贴 —— TSV(制表符+换行)粘到指定左上角,类型自动推断;undo 还原。
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: pasteText 2×2 区域写入 + 类型推断 + undo 还原`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    const before = await call(page, handle, 'getCellValue', 2, 1) // B3 原值(将被覆盖)

    // 粘 2 行 2 列到 (2,1):数字串→数字
    expect(await call(page, handle, 'pasteText', '11\t22\n33\t44', { row: 2, col: 1 })).toBe(true)
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBe(11)
    expect(await call(page, handle, 'getCellValue', 2, 2)).toBe(22)
    expect(await call(page, handle, 'getCellValue', 3, 1)).toBe(33)
    expect(await call(page, handle, 'getCellValue', 3, 2)).toBe(44)

    // undo → 一次撤销整块粘贴
    await call(page, handle, 'undo')
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBe(before)

    // 文本 + 公式推断
    await call(page, handle, 'pasteText', 'hi\t=1+2', { row: 4, col: 1 })
    expect(await call(page, handle, 'getCellValue', 4, 1)).toBe('hi') // 文本
    const snap = (await call(page, handle, 'getCellSnapshot', 4, 2)) as { cell: { type: string; formula?: string } }
    expect(snap.cell.type).toBe('formula') // = 开头 → 公式
  })
}

test.describe('粘贴 e2e(G2:TSV → 区域 + 类型推断 + undo)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
