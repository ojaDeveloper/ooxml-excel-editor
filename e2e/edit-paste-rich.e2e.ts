import { test, expect, type Page } from '@playwright/test'

async function ready(page: Page, canvasSel: string, handle: string) {
  await page.getByRole('button', { name: '加载示例' }).click()
  await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
  await page.waitForFunction((h) => (window as any)[h] != null, handle, { timeout: 20_000 })
  await page.getByText('编辑模式').click()
  await page.waitForFunction((h) => (window as any)[h].isCellEditable(2, 1), handle, { timeout: 5_000 })
}

const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

const HTML =
  '<table><tr>' +
  '<td style="font-weight:bold;background:#ffff00;color:#ff0000">7</td>' +
  '<td style="font-style:italic">8</td>' +
  '</tr><tr><td colspan="2" style="text-align:center">x</td></tr></table>'

// Phase C:富粘贴 —— Excel/WPS HTML → 值 + 字体/颜色/填充/对齐 + 合并,整体单次撤销
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 富粘贴 HTML(值+样式+合并)+ 单次撤销`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    const orig = await call(page, handle, 'getCellValue', 2, 1) // 粘贴前原值(样例里非空)
    // 落点 (2,1):row2 可编辑(demo 只读区是 row1)
    expect(await call(page, handle, 'pasteRichHtml', HTML, { row: 2, col: 1 })).toBe(true)

    // 值:数字推断
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBe(7)
    expect(await call(page, handle, 'getCellValue', 2, 2)).toBe(8)
    // 样式:粗体 + 黄填充 + 红字
    const snap = (await call(page, handle, 'getCellSnapshot', 2, 1)) as any
    expect(snap.style.font.bold).toBe(true)
    expect(String(snap.style.fill.fgColor).toUpperCase()).toBe('#FFFF00')
    expect(String(snap.style.font.color).toUpperCase()).toBe('#FF0000')
    // 斜体
    expect(((await call(page, handle, 'getCellSnapshot', 2, 2)) as any).style.font.italic).toBe(true)
    // 合并:colspan=2 → (3,1)-(3,2)
    const merges = (await page.evaluate((h) => {
      const v = (window as any)[h]
      return v.getWorkbook().sheets[v.getActiveSheet()].merges
    }, handle)) as { top: number; left: number; bottom: number; right: number }[]
    expect(merges.some((m) => m.top === 3 && m.left === 1 && m.bottom === 3 && m.right === 2)).toBe(true)

    // 单次撤销 → 值/样式/合并全部回退(值恢复成粘贴前原值)
    await call(page, handle, 'undo')
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBe(orig)
    const merges2 = (await page.evaluate((h) => {
      const v = (window as any)[h]
      return v.getWorkbook().sheets[v.getActiveSheet()].merges
    }, handle)) as { top: number; left: number }[]
    expect(merges2.some((m) => m.top === 3 && m.left === 1)).toBe(false)
  })
}

test.describe('富粘贴 e2e(Phase C:Excel/WPS HTML → 样式/合并 + 单次撤销)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
