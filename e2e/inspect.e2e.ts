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

const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

// inspectCell(row, col) —— 单元格全息体检:snapshot + 合并 + 浮动图 + DISPIMG + 数据验证 + 条件格式 + 链接/批注
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: inspectCell 合并锚点/被覆盖/普通格`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // 示例第一行整行合并(A1:E1)。
    const anchor = (await call(page, handle, 'inspectCell', 0, 0)) as any
    expect(anchor).not.toBeNull()
    expect(anchor.merge).not.toBeNull()
    expect(anchor.merge.top).toBe(0)
    expect(anchor.merge.left).toBe(0)
    expect(anchor.isMergeAnchor).toBe(true)

    const covered = (await call(page, handle, 'inspectCell', 0, 2)) as any
    expect(covered.merge).not.toBeNull()
    expect(covered.isMergeAnchor).toBe(false)

    const plain = (await call(page, handle, 'inspectCell', 2, 0)) as any
    expect(plain.merge).toBeNull()
    expect(plain.isMergeAnchor).toBe(false)
    expect(plain.floatingImages).toEqual([])
    expect(plain.cellImage).toBeNull()
    // snapshot 字段透传:有 raw/text
    expect(typeof plain.text).toBe('string')
  })
}

test.describe('Cell Inspector e2e(P1)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
