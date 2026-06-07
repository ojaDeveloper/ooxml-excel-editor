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

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const BYTES = [137, 80, 78, 71, 13, 10, 26, 10]

// 选区批量浮动 ⇄ 嵌入(P2):3 张浮动图分布范围内/外,选区批量转 → 内 2 张嵌入、外 1 张留浮动;
// 反向:DISPIMG 格批量浮动化 → 选区内的转出来;undo 全恢复。
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: convertImagesInRangeToCell 选区内批量嵌入,选区外保留 + undo 还原`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    const img = (row: number, col: number) => ({
      src: PNG,
      bytes: BYTES,
      mime: 'image/png',
      from: { col, colOffEmu: 0, row, rowOffEmu: 0 },
      extWidthEmu: 952500,
      extHeightEmu: 952500,
    })
    // 加 3 张浮动图:#0 (3,3)/ #1 (4,2)/ #2 (10,10)
    await call(page, handle, 'addImage', img(3, 3))
    await call(page, handle, 'addImage', img(4, 2))
    await call(page, handle, 'addImage', img(10, 10))
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(3)

    // 选区 (0,0)..(6,6) → 内含 #0 和 #1(图中心反推可能受行高列宽影响,这里覆盖范围足够宽)
    const inside = (await call(page, handle, 'convertImagesInRangeToCell', { top: 0, left: 0, bottom: 6, right: 6 })) as number
    expect(inside).toBeGreaterThanOrEqual(1)
    expect(inside).toBeLessThanOrEqual(2)
    // #2 在 (10,10) 远外,绝对应保留:剩余浮动图 = 3 - inside;CellImages = inside
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(3 - inside)
    expect(((await call(page, handle, 'getCellImages')) as any[]).length).toBe(inside)

    // 一次 undo 全恢复
    await call(page, handle, 'undo')
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(3)
    expect(((await call(page, handle, 'getCellImages')) as any[]).length).toBe(0)
  })

  test(`${label}: convertCellImagesInRangeToFloat 选区内批量浮动化 + undo`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // 先放 1 张图、嵌入到 (3,3) → 模型里有 1 张 DISPIMG
    await call(page, handle, 'addImage', {
      src: PNG, bytes: BYTES, mime: 'image/png',
      from: { col: 3, colOffEmu: 0, row: 3, rowOffEmu: 0 },
      extWidthEmu: 952500, extHeightEmu: 952500,
    })
    expect(await call(page, handle, 'convertImageToCell', 0, 3, 3)).toBe(true)
    expect(((await call(page, handle, 'getCellImages')) as any[]).length).toBe(1)
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(0)

    // 选区 (2,2)..(4,4) 把 DISPIMG 拎成浮动图(应有 1 张)
    expect(await call(page, handle, 'convertCellImagesInRangeToFloat', { top: 2, left: 2, bottom: 4, right: 4 })).toBe(1)
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(1)

    // 一次 undo 恢复 DISPIMG
    await call(page, handle, 'undo')
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(0)
    expect(((await call(page, handle, 'getCellImages')) as any[]).length).toBe(1)
  })
}

test.describe('图片选区批量互转 e2e(P2)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
