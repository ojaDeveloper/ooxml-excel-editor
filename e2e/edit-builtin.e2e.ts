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

async function cellCenter(page: Page, renderAreaSel: string, handle: string, row: number, col: number) {
  const box = (await page.locator(renderAreaSel).boundingBox())!
  const rect = (await call(page, handle, 'rectOf', row, col)) as { x: number; y: number; w: number; h: number }
  return { x: box.x + rect.x + rect.w / 2, y: box.y + rect.y + rect.h / 2 }
}

// E3:内置文本编辑器 —— 双击/打字进编辑,Enter 提交+下移,Esc 取消;只读格双击不进编辑。
function run(label: string, url: string, canvasSel: string, renderAreaSel: string, handle: string) {
  test(`${label}: 双击进编辑→改值 Enter 提交;打字进编辑;Esc 取消;只读不进`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    const input = page.locator('input.ooxml-cell-editor')

    // 双击单价格(2,1,无自定义编辑器 → 内置文本编辑器)
    const c21 = await cellCenter(page, renderAreaSel, handle, 2, 1)
    await page.mouse.dblclick(c21.x, c21.y)
    await expect(input).toBeVisible()

    // 编辑框精确贴合单元格(不放大):尺寸应约等于该格 rect(±1px 容差,排除取整)
    const rect21 = (await call(page, handle, 'rectOf', 2, 1)) as { w: number; h: number }
    const ib = (await input.boundingBox())!
    expect(Math.abs(ib.width - rect21.w)).toBeLessThanOrEqual(1)
    expect(Math.abs(ib.height - rect21.h)).toBeLessThanOrEqual(1)

    // 改 999 + Enter → 提交(保留货币格式)+ 编辑器关 + cell-change
    await input.fill('999')
    await input.press('Enter')
    await expect(input).toBeHidden()
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBe(999)
    const evt = await page.evaluate(() => (window as any).__lastCellChange)
    expect(evt.after.raw).toBe(999)

    // 打字进编辑:单击数量格(2,2)选中 → 按 '7' → 编辑器出现且值为 '7'
    const c22 = await cellCenter(page, renderAreaSel, handle, 2, 2)
    await page.mouse.click(c22.x, c22.y)
    await page.keyboard.press('7')
    await expect(input).toBeVisible()
    await expect(input).toHaveValue('7')
    // Esc 取消 → 编辑器关、值不变
    const before22 = await call(page, handle, 'getCellValue', 2, 2)
    await input.press('Escape')
    await expect(input).toBeHidden()
    expect(await call(page, handle, 'getCellValue', 2, 2)).toBe(before22)

    // 只读表头格(1,1)双击 → 不进编辑
    const c11 = await cellCenter(page, renderAreaSel, handle, 1, 1)
    await page.mouse.dblclick(c11.x, c11.y)
    await expect(input).toBeHidden()

    // 合并单元格(A1:E1)从被覆盖格(0,3)进入 → 落到锚点(0,0):
    // 锚点在第 0 列 → demo 给的是自定义 <select> 编辑器(顺带验证"覆盖格→锚点"重定向),
    // 且编辑框盖住整片合并区(宽 ≈ A..E 之和,不是单格)。
    const mergeEd = page.locator('select.demo-cell-editor')
    expect(await call(page, handle, 'beginEdit', 0, 3)).toBe(true)
    await expect(mergeEd).toBeVisible()
    const r00 = (await call(page, handle, 'rectOf', 0, 0)) as { w: number }
    let mergeW = 0
    for (let c = 0; c < 5; c++) mergeW += ((await call(page, handle, 'rectOf', 0, c)) as { w: number }).w
    const mb = (await mergeEd.boundingBox())!
    expect(mb.width).toBeGreaterThan(r00.w * 2) // 明显比单格宽
    expect(Math.abs(mb.width - mergeW)).toBeLessThanOrEqual(1) // ≈ 整片合并区宽
    await call(page, handle, 'cancelEdit')
    await expect(mergeEd).toBeHidden()
  })
}

test.describe('内置编辑器 e2e(E3:文本编辑 + 键盘进编辑)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '.render-area', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '.rxl-render-area', '__excelViewerReact')
})
