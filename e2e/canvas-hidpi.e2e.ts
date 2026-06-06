import { test, expect, type Page } from '@playwright/test'

// 模拟高 DPI 屏 / 系统缩放(devicePixelRatio≠1)。浏览器 Ctrl+缩放、Windows 125%/150% 都会让 dpr 变成分数。
test.use({ deviceScaleFactor: 1.5 })

async function ready(page: Page, canvasSel: string, handle: string) {
  await page.getByRole('button', { name: '加载示例' }).click()
  await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
  await page.waitForFunction((h) => (window as any)[h] != null, handle, { timeout: 20_000 })
  // 等网格真正绘制(canvas 缓冲已按 dpr 调整,不再是默认 300×150)
  await page.waitForFunction(
    (sel) => {
      const cv = document.querySelector(sel) as HTMLCanvasElement | null
      return !!(cv && cv.width > 400)
    },
    canvasSel,
    { timeout: 20_000 },
  )
}
const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

// dpr≠1 时,canvas(替换元素)若不显式钉死 CSS 显示尺寸,会以缓冲像素(width*dpr)显示 → 比容器大 dpr 倍 →
// 网格放大、与 DOM 叠加层(浮动图/图表)及鼠标命中错位。这里锁两条不变量:
//   ① canvas 屏幕显示尺寸 == 容器(网格逻辑坐标空间)尺寸;② 在单元格视觉中心点击命中该格。
function run(label: string, url: string, canvasSel: string, handle: string, renderAreaSel: string) {
  test(`${label}: dpr=1.5 下 canvas 显示尺寸=容器 + 点击命中`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // ① canvas 显示尺寸 == 容器逻辑尺寸(buffer 仍是 dpr 倍,只是显示被降采样)
    const g = (await page.evaluate(
      ([raSel, cvSel]) => {
        const ra = document.querySelector(raSel) as HTMLElement
        const cv = document.querySelector(cvSel) as HTMLCanvasElement
        const cr = cv.getBoundingClientRect()
        return { dpr: window.devicePixelRatio, raW: ra.clientWidth, raH: ra.clientHeight, bufW: cv.width, dispW: cr.width, dispH: cr.height }
      },
      [renderAreaSel, canvasSel],
    )) as { dpr: number; raW: number; raH: number; bufW: number; dispW: number; dispH: number }
    expect(g.dpr).toBeCloseTo(1.5, 1)
    expect(g.bufW).toBeGreaterThan(g.raW) // 缓冲是高清的(> 逻辑宽)
    expect(Math.abs(g.dispW - g.raW)).toBeLessThan(2) // 但显示宽 == 逻辑宽
    expect(Math.abs(g.dispH - g.raH)).toBeLessThan(2)

    // ② 在单元格"视觉中心"点击应命中该格(canvas 与逻辑 1:1 → box+rectOf 即视觉位置)
    for (const t of [{ row: 1, col: 1 }, { row: 7, col: 3 }]) {
      await call(page, handle, 'setSelection', { top: 0, left: 0, bottom: 0, right: 0 })
      const r = (await call(page, handle, 'rectOf', t.row, t.col)) as { x: number; y: number; w: number; h: number }
      const box = (await page.locator(renderAreaSel).boundingBox())!
      await page.mouse.click(box.x + r.x + r.w / 2, box.y + r.y + r.h / 2)
      const sel = (await call(page, handle, 'getSelection')) as { top: number; left: number }
      expect(`${sel.top},${sel.left}`, `${label} cell ${t.row},${t.col}`).toBe(`${t.row},${t.col}`)
    }
  })
}

test.describe('HiDPI / 系统缩放 canvas 对齐 e2e(dpr≠1)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer', '.render-area')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact', '.rxl-render-area')
})
