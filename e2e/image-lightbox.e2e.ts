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
  await page.waitForFunction((h) => (window as any)[h] != null, handle, { timeout: 20_000 })
}

const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const BYTES = [137, 80, 78, 71, 13, 10, 26, 10]

// Phase B:图片点击放大灯箱 + 下载;编辑模式右键内嵌图格「查看大图」弹大图
function run(label: string, url: string, canvasSel: string, handle: string, renderAreaSel: string) {
  test(`${label}: 命令式灯箱 + 下载 + 右键内嵌图查看大图`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    const lb = page.locator('.ooxml-lightbox')
    const menu = page.locator('.ooxml-context-menu')

    // 1) 命令式打开灯箱 → 大图 + 下载按钮;Esc 关闭
    await call(page, handle, 'openImageLightbox', PNG, 'pic.png', 'image/png')
    await expect(lb).toBeVisible()
    await expect(lb.locator('img')).toBeVisible()
    await expect(lb.getByText('下载原图')).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(lb).toBeHidden()

    // 2) 编辑模式:加带字节浮动图 → 转为内嵌图格 (5,1)
    await page.getByText('编辑模式').click()
    await page.waitForFunction((h) => (window as any)[h].isCellEditable(5, 1), handle, { timeout: 5_000 })
    await call(page, handle, 'addImage', { src: PNG, bytes: BYTES, mime: 'image/png', from: { col: 1, colOffEmu: 0, row: 5, rowOffEmu: 0 }, extWidthEmu: 800000, extHeightEmu: 800000 })
    expect(await call(page, handle, 'convertImageToCell', 0, 5, 1)).toBe(true)
    expect(await call(page, handle, 'getCellImageAt', 5, 1)).not.toBeNull() // (5,1) 现是内嵌图

    // 3) 右键内嵌图格 → 菜单「查看大图 / 下载原图」→ 点开灯箱
    const box = (await page.locator(renderAreaSel).boundingBox())!
    const r = (await call(page, handle, 'rectOf', 5, 1)) as { x: number; y: number; w: number; h: number }
    await page.mouse.click(box.x + r.x + r.w / 2, box.y + r.y + r.h / 2, { button: 'right' })
    await expect(menu).toBeVisible()
    await menu.getByText('查看大图').click()
    await expect(lb).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(lb).toBeHidden()
  })
}

test.describe('图片点击放大灯箱 + 下载 e2e(Phase B)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer', '.render-area')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact', '.rxl-render-area')
})
