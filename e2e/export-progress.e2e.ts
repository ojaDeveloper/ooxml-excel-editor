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

// 导出进度 + 取消:
//   ① exportImage 订阅 onProgress → 计数 ≥ 1(至少 render+write 两次)
//   ② exportImage 传入 pre-aborted signal → 抛 AbortError 而非完成
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: onProgress 计数 > 0 + pre-aborted signal 抛 AbortError`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // ① onProgress 在 exportImage 路径里 emit
    const progressStages = await page.evaluate(async (h) => {
      const stages: string[] = []
      const blob = await (window as any)[h].exportImage({
        onProgress: (p: { stage: string }) => stages.push(p.stage),
      })
      return { stages, hasBlob: blob && (blob as Blob).size > 0 }
    }, handle)
    expect(progressStages.hasBlob).toBe(true)
    expect(progressStages.stages.length).toBeGreaterThan(0)
    // 至少有一次 'render' 和一次 'write'(开始 + 结束)
    expect(progressStages.stages).toContain('render')
    expect(progressStages.stages).toContain('write')

    // ② pre-aborted signal:exportImage 应抛 AbortError(不产 blob)
    const aborted = await page.evaluate(async (h) => {
      const ctrl = new AbortController()
      ctrl.abort()
      try {
        await (window as any)[h].exportImage({ signal: ctrl.signal })
        return { ok: false, errName: null as string | null }
      } catch (e) {
        return { ok: true, errName: (e as Error & { name: string }).name }
      }
    }, handle)
    expect(aborted.ok).toBe(true)
    expect(aborted.errName).toBe('AbortError')
  })
}

test.describe('导出进度 + 取消(P1)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
