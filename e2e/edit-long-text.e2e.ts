import { test, expect, type Page } from '@playwright/test'

/**
 * WPS 风格长文本编辑 e2e (Phase 1, 2026-06-08).
 *
 * 验证:
 *   ① 编辑器 root 是 <textarea> (Phase 1 改的)
 *   ② 短文本: 编辑器高度 ≈ 单元格原高 (不撑大)
 *   ③ 长文本: 编辑器高度 > 单元格原高, 向下溢出原格 (textarea 多行换行)
 *   ④ 输入新文本: 高度跟着内容动态变化
 *   ⑤ Esc 取消: 行高恢复原始 (本来就没改, 验证 happy path)
 *   ⑥ Shift+Enter 插入换行不提交; 普通 Enter 提交
 */

async function loadSampleEditable(page: Page, url: string, canvasSel: string, handle: string) {
  await page.goto(url)
  await page.getByRole('button', { name: '加载示例' }).click()
  await page.waitForFunction((h) => (window as any)[h]?.getWorkbook?.() != null, handle, { timeout: 20_000 })
  await page.locator(canvasSel).waitFor({ state: 'visible', timeout: 20_000 })
  await page.getByText('编辑模式').click()
}

const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: ① 编辑器 root 是 textarea`, async ({ page }) => {
    await loadSampleEditable(page, url, canvasSel, handle)
    // 命令式进入编辑 A1
    const dbg = await page.evaluate((h: string) => {
      const v = (window as any)[h]
      return {
        editable: v.isCellEditable(2, 1),
        editingBefore: v.isEditing(),
        canBegin: typeof v.beginEdit,
        beginResult: v.beginEdit(2, 1),
        editingAfter: v.isEditing(),
      }
    }, handle)
    expect(dbg.canBegin).toBe('function')
    expect(dbg.editable).toBe(true)
    expect(dbg.beginResult).toBe(true)
    expect(dbg.editingAfter).toBe(true)
    // 编辑器层应该有一个 textarea (Phase 1 改的)
    const tag = await page.locator('.ooxml-cell-editor').evaluate((el) => el.tagName)
    expect(tag).toBe('TEXTAREA')
  })

  test(`${label}: ② 短文本编辑器高度 ≈ 单元格原高`, async ({ page }) => {
    await loadSampleEditable(page, url, canvasSel, handle)
    // 取 A1 的原始矩形高度
    const cellH = await page.evaluate((h: string) => {
      const v = (window as any)[h]
      return (v.rectOf(2, 1) as { h: number }).h
    }, handle)
    await call(page, handle, 'beginEdit', 2, 1)
    const editorH = await page.locator('.ooxml-cell-editor').evaluate((el) => parseFloat((el as HTMLElement).style.height))
    // 应该接近 cell 高 (短文本不大幅撑大, 允许 +12 px 的边框 / 行盒细微差异)
    expect(editorH).toBeLessThanOrEqual(cellH + 12)
  })

  test(`${label}: ③ 长文本编辑器向下撑高溢出原格`, async ({ page }) => {
    await loadSampleEditable(page, url, canvasSel, handle)
    const cellH = await page.evaluate((h: string) => {
      const v = (window as any)[h]
      return (v.rectOf(2, 1) as { h: number }).h
    }, handle)
    await call(page, handle, 'beginEdit', 2, 1)
    // 在 textarea 里输入超长文本 (复制截图里 WPS 那段)
    const ta = page.locator('.ooxml-cell-editor')
    const longText = '贴标ONEVAN 10个1500 英规2电1充 无刷牧田蓝色700NM扳手 机身+电池+塑箱+扳手套筒配件（sku001697） 牧田蓝色 1/2轴 实紧固700NM拆卸1000NM 电机 5225（电池贴ONEVAN通用标 4.0Ah）'
    await ta.fill(longText)
    // 等 input 事件 → reposition → host 撑高
    await page.waitForTimeout(50)
    const editorH = await ta.evaluate((el) => parseFloat((el as HTMLElement).style.height))
    expect(editorH).toBeGreaterThan(cellH * 2) // 至少撑出 2 倍原高
  })

  test(`${label}: ④ 输入中编辑器高度随内容动态变化`, async ({ page }) => {
    await loadSampleEditable(page, url, canvasSel, handle)
    await call(page, handle, 'beginEdit', 2, 1)
    const ta = page.locator('.ooxml-cell-editor')
    await ta.fill('短')
    await page.waitForTimeout(30)
    const h1 = await ta.evaluate((el) => parseFloat((el as HTMLElement).style.height))
    // 加 5 个 \n 显式换行
    await ta.fill('line1\nline2\nline3\nline4\nline5\nline6')
    await page.waitForTimeout(30)
    const h2 = await ta.evaluate((el) => parseFloat((el as HTMLElement).style.height))
    expect(h2).toBeGreaterThan(h1) // 内容多了, 高度涨
  })

  test(`${label}: ⑤ Esc 取消, 编辑器卸载 + 单元格 raw 不变`, async ({ page }) => {
    await loadSampleEditable(page, url, canvasSel, handle)
    const beforeRaw = await call(page, handle, 'getCellValue', 2, 1)
    await call(page, handle, 'beginEdit', 2, 1)
    const ta = page.locator('.ooxml-cell-editor')
    await ta.fill('CANCEL_ME')
    await ta.press('Escape')
    await expect(ta).toBeHidden({ timeout: 1000 })
    const afterRaw = await call(page, handle, 'getCellValue', 2, 1)
    expect(afterRaw).toEqual(beforeRaw) // 取消后值不变
  })

  test(`${label}: ⑥ Shift+Enter 插入换行不提交; 普通 Enter 提交`, async ({ page }) => {
    await loadSampleEditable(page, url, canvasSel, handle)
    await call(page, handle, 'beginEdit', 2, 1)
    const ta = page.locator('.ooxml-cell-editor')
    await ta.fill('line1')
    await ta.press('Shift+Enter')
    // 编辑器应仍存在 (没提交)
    await expect(ta).toBeVisible()
    // 验有 \n 插入 (textarea 当前 value 含换行)
    const val = await ta.inputValue()
    expect(val).toContain('\n')
    // 普通 Enter → 提交 + 卸载
    await ta.press('Enter')
    await expect(ta).toBeHidden({ timeout: 1000 })
  })
}

test.describe('WPS 长文本编辑 e2e (Phase 1, 2026-06-08)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
