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
  await page.getByText('编辑模式').click() // 开编辑
}

/** 命令式 editCell → 模型变更 + cell-change 前后快照 + undo/redo + 只读拦截 */
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: editCell → 模型变更 + cell-change 前后快照 + undo/redo + 只读拦截`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // 数据格(2,0)= "笔记本电脑" → 编辑为 "改过了"
    const r = await page.evaluate((h) => {
      const v = (window as any)[h]
      const ok = v.editCell(2, 0, '改过了')
      return { ok, text: v.getCellText(2, 0), snap: v.getCellSnapshot(2, 0) }
    }, handle)
    expect(r.ok).toBe(true)
    expect(r.text).toBe('改过了')
    expect(r.snap.raw).toBe('改过了')

    // cell-change 事件:前 "笔记本电脑" → 后 "改过了",source=api
    const evt = await page.evaluate(() => (window as any).__lastCellChange)
    expect(evt.source).toBe('api')
    expect(evt.before.text).toBe('笔记本电脑')
    expect(evt.after.text).toBe('改过了')

    // undo 还原
    const afterUndo = await page.evaluate((h) => {
      const v = (window as any)[h]
      v.undo()
      return v.getCellText(2, 0)
    }, handle)
    expect(afterUndo).toBe('笔记本电脑')

    // redo 重做
    const afterRedo = await page.evaluate((h) => {
      const v = (window as any)[h]
      v.redo()
      return v.getCellText(2, 0)
    }, handle)
    expect(afterRedo).toBe('改过了')

    // 只读格(表头行 1,在 readOnlyRanges)→ editCell 不生效
    const ro = await page.evaluate((h) => {
      const v = (window as any)[h]
      const ok = v.editCell(1, 0, 'X')
      return { ok, text: v.getCellText(1, 0) }
    }, handle)
    expect(ro.ok).toBe(false)
    expect(ro.text).not.toBe('X')
  })
}

test.describe('编辑 API e2e(E1:命令式编辑 + 快照事件 + 撤销重做)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
