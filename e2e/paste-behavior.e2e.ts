import { test, expect, type Page } from '@playwright/test'

async function ready(page: Page, canvasSel: string, handle: string) {
  await page.getByRole('button', { name: '加载示例' }).click()
  await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
  await page.waitForFunction((h) => (window as any)[h] != null, handle, { timeout: 20_000 })
  await page.getByText('编辑模式').click()
  await page.waitForFunction((h) => (window as any)[h].isCellEditable(5, 5), handle, { timeout: 5_000 })
}
const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

const VALUES_ONLY = { cellStyle: 'skip', fill: 'skip', rowHeight: 'keep', colWidth: 'keep', sourceMerges: 'skip', targetMerges: 'keep', images: 'skip' }

// 选择性粘贴「保留原样式(仅值)」:目标红底 + 源白底蓝字 → 仅填值,目标样式保留
function runValuesOnly(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 选择性粘贴「仅值」→ 保留目标样式、只填值`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    await call(page, handle, 'setStyle', { top: 5, left: 5, bottom: 5, right: 5 }, { fill: { type: 'solid', fgColor: '#FF0000' } })
    const HTML = '<table><tr><td style="background:#FFFFFF;color:#0000FF;font-weight:bold">X</td></tr></table>'
    expect(await call(page, handle, 'pasteRichHtml', HTML, { row: 5, col: 5 }, VALUES_ONLY)).toBe(true)
    expect(await call(page, handle, 'getCellValue', 5, 5)).toBe('X') // 值进来了
    const s = (await call(page, handle, 'getCellSnapshot', 5, 5)) as any
    expect(String(s.style.fill.fgColor).toUpperCase()).toBe('#FF0000') // 目标红底保留(没被源白底覆盖)
    expect(s.style.font.bold).toBe(false) // 源 bold 没覆盖(仅值)
  })
}

// 默认覆盖粘贴 → 清掉目标区原有合并(否则旧合并吞列,即截图那个 bug)
function runMergeClear(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 默认覆盖粘贴 → 清掉目标区原有合并(数据不再被旧合并吞列)`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    await call(page, handle, 'mergeCells', { top: 5, left: 5, bottom: 5, right: 7 })
    const merged = (await page.evaluate((h) => {
      const v = (window as any)[h]
      return v.getWorkbook().sheets[v.getActiveSheet()].merges.some((m: any) => m.top === 5 && m.left === 5 && m.right === 7)
    }, handle)) as boolean
    expect(merged).toBe(true)
    // 粘 3 列(默认覆盖,targetMerges=clear)
    expect(await call(page, handle, 'pasteRichHtml', '<table><tr><td>a</td><td>b</td><td>c</td></tr></table>', { row: 5, col: 5 })).toBe(true)
    const stillMerged = (await page.evaluate((h) => {
      const v = (window as any)[h]
      return v.getWorkbook().sheets[v.getActiveSheet()].merges.some((m: any) => m.top === 5 && m.left === 5)
    }, handle)) as boolean
    expect(stillMerged).toBe(false) // 旧合并清了
    expect(await call(page, handle, 'getCellValue', 5, 6)).toBe('b') // 被吞的中间列恢复
  })
}

// 工具栏「⚙ 粘贴配置」/ openPasteConfigDialog → 弹框架无关配置面板(7 项下拉)
function runConfigDialog(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: openPasteConfigDialog 弹粘贴配置面板(7 项 + 预设)`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    expect(await call(page, handle, 'openPasteConfigDialog')).toBe(true)
    await expect(page.locator('.ooxml-paste-config-mask')).toBeVisible()
    expect(await page.locator('.ooxml-paste-config-mask select[data-key]').count()).toBe(7)
    // 点「仅值」预设 → cellStyle 下拉变 skip
    await page.locator('.ooxml-paste-config-mask [data-preset-values]').click()
    expect(await page.locator('.ooxml-paste-config-mask select[data-key="cellStyle"]').inputValue()).toBe('skip')
    // 应用 → 关闭 + setPasteBehavior 生效
    await page.locator('.ooxml-paste-config-mask [data-ok]').click()
    await expect(page.locator('.ooxml-paste-config-mask')).toHaveCount(0)
    expect((await call(page, handle, 'getPasteBehavior') as any).cellStyle).toBe('skip')
  })
}

// 粘贴撞只读格(demo A2:E2 = readOnlyRanges)→ 默认弹「只读格未被覆盖」对话框列出具体格 + 那些格没被改
function runReadOnlyDialog(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 粘到只读区 → 弹只读提醒对话框(列出哪些格)+ 只读格不被覆盖`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)
    // demo 把 A2:E2(row 1, col 0-4)设了 readOnlyRanges;粘 3 格到 (1,0) 全撞只读
    const before = await call(page, handle, 'getCellValue', 1, 0) // 示例原值(只读)
    expect(await call(page, handle, 'pasteRichHtml', '<table><tr><td>X</td><td>Y</td><td>Z</td></tr></table>', { row: 1, col: 0 })).toBe(true)
    await expect(page.locator('.ooxml-readonly-mask')).toBeVisible() // 默认 dialog 弹出
    await expect(page.locator('.ooxml-readonly-mask')).toContainText('A2') // 列出具体只读格
    expect(await call(page, handle, 'getCellValue', 1, 0)).toBe(before) // 只读格没被覆盖(还是原值)
    await page.locator('.ooxml-readonly-mask [data-ok]').click()
    await expect(page.locator('.ooxml-readonly-mask')).toHaveCount(0)
  })
}

test.describe('粘贴行为配置 e2e(选择性粘贴 / 目标合并清除 / 配置面板)', () => {
  runReadOnlyDialog('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  runReadOnlyDialog('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
  runValuesOnly('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  runValuesOnly('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
  runMergeClear('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  runMergeClear('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
  runConfigDialog('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  runConfigDialog('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
