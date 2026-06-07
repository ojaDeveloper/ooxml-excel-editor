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
}

// 内置导出进度遮罩(P1.5):exportImage 时遮罩自动出现,完事自动消失;用户传 onProgress 仍被链回调
function run(label: string, url: string, canvasSel: string, overlaySel: string, handle: string) {
  test(`${label}: exportImage 期间 .export-progress-overlay / .rxl-export-progress 出现 + 用户 onProgress 仍被调`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // 监听遮罩出现:在 exportImage 调用前订阅 MutationObserver,记录第一次看到 overlay 的时机
    await page.evaluate(() => {
      ;(window as any).__overlaySeen = false
      const mo = new MutationObserver(() => {
        if (document.querySelector('.export-progress-overlay, .rxl-export-progress')) (window as any).__overlaySeen = true
      })
      mo.observe(document.body, { childList: true, subtree: true })
    })

    const r = await page.evaluate(async (h) => {
      const stages: string[] = []
      const blob = await (window as any)[h].exportImage({ onProgress: (p: { stage: string }) => stages.push(p.stage) })
      return { stages, size: (blob as Blob).size, seen: (window as any).__overlaySeen as boolean }
    }, handle)
    expect(r.size).toBeGreaterThan(0)
    expect(r.stages.length).toBeGreaterThan(0)
    expect(r.seen).toBe(true)
    // 导出完成,遮罩消失
    await expect(page.locator(overlaySel)).toHaveCount(0)
  })
}

test.describe('内置导出进度遮罩 e2e(P1.5)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.export-progress-overlay', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-export-progress', '__excelViewerReact')
})
