import { test, expect, type Page } from '@playwright/test'
import { promises as fs } from 'node:fs'
import path from 'node:path'

const OUT = path.join('test-results', 'downloads')

/** 加载示例并等待真实渲染到 canvas(像素非空) */
async function loadSample(page: Page) {
  await page.goto('/')
  await page.getByRole('button', { name: '加载示例' }).click()
  await expect(page.locator('canvas.grid-canvas')).toBeVisible({ timeout: 20_000 })
  // 等 worker 解析 + 渲染完成: canvas 上出现非白像素
  await page.waitForFunction(
    () => {
      const c = document.querySelector('canvas.grid-canvas') as HTMLCanvasElement | null
      if (!c || !c.width || !c.height) return false
      const ctx = c.getContext('2d')
      if (!ctx) return false
      const w = Math.min(c.width, 300)
      const h = Math.min(c.height, 300)
      const d = ctx.getImageData(0, 0, w, h).data
      for (let i = 0; i < d.length; i += 4) {
        if (d[i] < 248 || d[i + 1] < 248 || d[i + 2] < 248) return true
      }
      return false
    },
    null,
    { timeout: 20_000 },
  )
}

/** 触发一次下载并把字节读回来 */
async function captureDownload(page: Page, trigger: () => Promise<void>): Promise<Buffer> {
  const pending = page.waitForEvent('download', { timeout: 30_000 })
  await trigger()
  const dl = await pending
  await fs.mkdir(OUT, { recursive: true })
  const file = path.join(OUT, dl.suggestedFilename())
  await dl.saveAs(file)
  return fs.readFile(file)
}

async function openExportMenu(page: Page) {
  await page.locator('.export-btn').click()
}

test.describe('导出 e2e(真浏览器: 解析→canvas 渲染→jsPDF→下载)', () => {
  test('加载示例 → 真实渲染到 canvas(非空像素)', async ({ page }) => {
    await loadSample(page) // 内部已断言 canvas 非空
    await expect(page.locator('.formula-bar')).toBeVisible()
  })

  test('PNG 导出 → 有效 PNG 文件', async ({ page }) => {
    await loadSample(page)
    const buf = await captureDownload(page, async () => {
      await openExportMenu(page)
      await page.getByRole('button', { name: '导出为图片 (PNG)' }).click()
    })
    expect(buf.length).toBeGreaterThan(2000)
    // PNG 魔数: 89 50 4E 47
    expect(buf[0]).toBe(0x89)
    expect(buf.subarray(1, 4).toString('latin1')).toBe('PNG')
  })

  test('位图 PDF → %PDF,整表贴图(含图像 XObject)', async ({ page }) => {
    await loadSample(page)
    const buf = await captureDownload(page, async () => {
      await openExportMenu(page)
      await page.getByRole('button', { name: '导出为 PDF (位图)' }).click()
    })
    expect(buf.subarray(0, 4).toString('latin1')).toBe('%PDF')
    expect(buf.length).toBeGreaterThan(3000)
    expect(buf.toString('latin1').includes('/Image')).toBeTruthy() // 整表贴图
  })

  // 用文字显示操作符(Tj)的数量区分: 矢量逐格写真文字 → 大量 Tj;位图整表贴图 → 几乎没有。
  // (/BaseFont 两种都有——jsPDF 总会写标准字体;Tj 计数才是真正的"有没有可选文字"信号)
  test('矢量 PDF → %PDF,且文字操作符明显多于位图(可选可搜)', async ({ page }) => {
    await loadSample(page)
    const countTj = (b: Buffer) => (b.toString('latin1').match(/Tj/g) || []).length
    const bitmap = await captureDownload(page, async () => {
      await openExportMenu(page)
      await page.getByRole('button', { name: '导出为 PDF (位图)' }).click()
    })
    const vector = await captureDownload(page, async () => {
      await openExportMenu(page)
      await page.getByRole('button', { name: '导出为 PDF (矢量·文字可选)' }).click()
    })
    expect(vector.subarray(0, 4).toString('latin1')).toBe('%PDF')
    expect(countTj(vector)).toBeGreaterThan(countTj(bitmap) + 5)
  })

  test('导出设置对话框 → 当前工作表导出 PDF', async ({ page }) => {
    await loadSample(page)
    await openExportMenu(page)
    await page.getByRole('button', { name: '导出设置…' }).click()
    await expect(page.getByText('导出 / 打印设置')).toBeVisible()
    // 选"当前工作表"
    await page.getByText('当前工作表', { exact: false }).click()
    const buf = await captureDownload(page, async () => {
      await page.getByRole('button', { name: '导出 PDF', exact: true }).click()
    })
    expect(buf.subarray(0, 4).toString('latin1')).toBe('%PDF')
  })
})
