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

// 1x1 透明 png(data url),给 addImage 一个合法 src(无需 blob)
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

// E6:图片编辑 —— addImage 加图 → moveImage 移动(归一为原点锚点)→ undo 还原;removeImage → undo 复原。
function run(label: string, url: string, canvasSel: string, handle: string, renderAreaSel: string) {
  test(`${label}: 鼠标拖拽浮动图 → 移动 + isDirty + undo 复位`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    // 加一张 100x100 图,锚在 (1,1) 偏移 0 → 屏幕左上角 = rectOf(1,1)
    await call(page, handle, 'addImage', {
      src: PNG,
      from: { col: 1, colOffEmu: 0, row: 1, rowOffEmu: 0 },
      extWidthEmu: 952500,
      extHeightEmu: 952500,
    })
    const box = (await page.locator(renderAreaSel).boundingBox())!
    const r = (await call(page, handle, 'rectOf', 1, 1)) as { x: number; y: number }
    const sx = box.x + r.x + 20 // 图内一点(避开边缘)
    const sy = box.y + r.y + 20
    await page.mouse.move(sx, sy)
    await page.mouse.down()
    await page.mouse.move(sx + 70, sy + 50, { steps: 5 })
    await page.mouse.up()

    const imgs = (await call(page, handle, 'getImages')) as any[]
    expect(imgs[0].from.col).toBe(0) // 拖拽命中 + setImageRect 归一
    expect(await call(page, handle, 'isDirty')).toBe(true)
    await call(page, handle, 'undo')
    expect(((await call(page, handle, 'getImages')) as any[])[0].from).toMatchObject({ col: 1, colOffEmu: 0 })
  })

  test(`${label}: addImage → moveImage + image-change + undo 还原;removeImage + undo 复原`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    const anchor = { src: PNG, from: { col: 1, colOffEmu: 0, row: 1, rowOffEmu: 0 }, extWidthEmu: 952500, extHeightEmu: 952500 }
    const idx = (await call(page, handle, 'addImage', anchor)) as number
    expect(idx).toBe(0)
    let imgs = (await call(page, handle, 'getImages')) as any[]
    expect(imgs).toHaveLength(1)
    expect((await win(page, '__lastImageChange')).before).toBeNull() // 加图:before null

    // 移动 +60,+40 → 归一为原点锚点(from.col=0),宽不变
    expect(await call(page, handle, 'moveImage', 0, 60, 40)).toBe(true)
    imgs = (await call(page, handle, 'getImages')) as any[]
    expect(imgs[0].from.col).toBe(0) // 已归一为原点相对
    expect(imgs[0].extWidthEmu).toBe(952500) // 移动不改尺寸
    const moveEvt = await win(page, '__lastImageChange')
    expect(moveEvt.before.from.col).toBe(1) // 前态(加图时的锚点)
    expect(moveEvt.after.from.col).toBe(0)

    // undo 移动 → 精确还原前态锚点(from.col 回 1,colOffEmu 回 0)
    await call(page, handle, 'undo')
    imgs = (await call(page, handle, 'getImages')) as any[]
    expect(imgs[0].from).toMatchObject({ col: 1, colOffEmu: 0 })

    // 删图 → undo 复原
    expect(await call(page, handle, 'removeImage', 0)).toBe(true)
    expect((await call(page, handle, 'getImages')) as any[]).toHaveLength(0)
    await call(page, handle, 'undo')
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(1)
  })
}

test.describe('图片编辑 e2e(E6:增删移改 + image-change + undo)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer', '.render-area')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact', '.rxl-render-area')
})
