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
// data url 没 bytes → 转换需要 bytes;给一张带 bytes 的图(1x1 透明 png 的字节)
const BYTES = [137, 80, 78, 71, 13, 10, 26, 10] // 仅占位,convert 只看 bytes 非空

// 第二期:浮动图 ⇄ WPS 单元格内嵌图(DISPIMG)一键互转 + undo/redo
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 浮动图 → 嵌入格 → 浮动图 往返 + undo`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // 加一张带字节的浮动图(转换要求 bytes/mime)
    await call(page, handle, 'addImage', {
      src: PNG,
      bytes: BYTES, // 转换只看 bytes 非空(普通数组即可,免去跨 bridge 的 TypedArray 序列化)
      mime: 'image/png',
      from: { col: 1, colOffEmu: 0, row: 5, rowOffEmu: 0 },
      extWidthEmu: 952500,
      extHeightEmu: 952500,
    })
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(1)
    expect(((await call(page, handle, 'getCellImages')) as any[]).length).toBe(0)

    // 浮动图 → 嵌入 (3,3)
    expect(await call(page, handle, 'convertImageToCell', 0, 3, 3)).toBe(true)
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(0)
    const reg = (await call(page, handle, 'getCellImages')) as any[]
    expect(reg.length).toBe(1)
    const snap = (await call(page, handle, 'getCellSnapshot', 3, 3)) as any
    expect(snap.cell.dispImgId).toBe(reg[0].id)
    expect(await call(page, handle, 'isDirty')).toBe(true)

    // undo → 回到浮动图
    await call(page, handle, 'undo')
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(1)
    expect(((await call(page, handle, 'getCellImages')) as any[]).length).toBe(0)

    // redo → 再次嵌入
    await call(page, handle, 'redo')
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(0)

    // 嵌入 → 浮动(拎回来)
    expect(await call(page, handle, 'convertCellImageToFloat', 3, 3)).toBe(true)
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(1)
    expect(((await call(page, handle, 'getCellImages')) as any[]).length).toBe(0) // 无引用 → 回收
    expect((await call(page, handle, 'getCellSnapshot', 3, 3) as any)?.cell).toBeNull()
  })

  test(`${label}: 就近批量嵌入(整表)+ 单次 undo + 贴合方式切换`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    // 加两张带字节的浮动图
    const anchor = (r: number) => ({ src: PNG, bytes: BYTES, mime: 'image/png', from: { col: 1, colOffEmu: 0, row: r, rowOffEmu: 0 }, extWidthEmu: 600000, extHeightEmu: 600000 })
    await call(page, handle, 'addImage', anchor(4))
    await call(page, handle, 'addImage', anchor(8))
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(2)

    // 整表就近嵌入:两张都进各自单元格,一次入栈
    const n = (await call(page, handle, 'convertAllImagesToCells')) as number
    expect(n).toBe(2)
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(0)
    expect(((await call(page, handle, 'getCellImages')) as any[]).length).toBe(2)

    // 贴合方式切换(不报错、即时重绘)
    await call(page, handle, 'setCellImageFit', 'contain')
    await call(page, handle, 'setCellImageFit', 'cover')
    await call(page, handle, 'setCellImageFit', 'fill')

    // 单次 undo 撤回整批
    await call(page, handle, 'undo')
    expect(((await call(page, handle, 'getImages')) as any[]).length).toBe(2)
    expect(((await call(page, handle, 'getCellImages')) as any[]).length).toBe(0)
  })
}

test.describe('WPS 单元格内嵌图 ⇄ 浮动图互转 e2e(第二期)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
