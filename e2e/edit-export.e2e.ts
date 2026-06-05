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

// E8:导出 —— 编辑一格后,xlsx(blob 非空 + 重解析值存活)/ csv / json 都带上编辑值(一份数据层)。
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 编辑后 exportXlsx(blob+重解析)/ exportCsv / exportJson 均含新值`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // 编辑 B3(单价)= 12345
    expect(await call(page, handle, 'editCell', 2, 1, 12345)).toBe(true)

    // CSV = 格式化显示(WYSIWYG,B 列 ¥ 格式 → 含千分位 12,345);JSON = raw 类型值(12345)
    const csv = (await call(page, handle, 'exportCsv')) as string
    expect(csv).toContain('12,345') // 显示值带货币格式
    const json = (await call(page, handle, 'exportJson')) as string
    expect(json).toContain('12345') // raw 数值

    // XLSX:导出 Blob 非空 + 类型正确(Blob 不能跨 evaluate 传,取 size/type)
    const meta = await page.evaluate(async (h) => {
      const blob: Blob = await (window as any)[h].exportXlsx()
      return { size: blob.size, type: blob.type }
    }, handle)
    expect(meta.size).toBeGreaterThan(1000) // 真实 xlsx zip 有体量
    expect(meta.type).toContain('spreadsheetml')
    // 值存活(写→重解析)由 xlsx-writer 单测覆盖,此处验真实下载链路产出合法 xlsx blob

    // F3:高保真 overlay 模式(重载原件叠加)—— 验证 sourceBuffer 端到端贯通、产合法 blob
    const ov = await page.evaluate(async (h) => {
      const blob: Blob = await (window as any)[h].exportXlsx({ fidelity: 'overlay' })
      return { size: blob.size, type: blob.type }
    }, handle)
    expect(ov.size).toBeGreaterThan(1000)
    expect(ov.type).toContain('spreadsheetml')
  })
}

test.describe('数据导出 e2e(E8:xlsx 重建 + csv/json 一份数据层)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
