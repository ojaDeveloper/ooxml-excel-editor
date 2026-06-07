import { test, expect, type Page } from '@playwright/test'

// JSON 直渲 + 模板填值(P3)。绕过 :src 加载示例 .xlsx 的常规路径,直接调命令式 API:
//   - jsonToWorkbook(...) → 通过 viewerApi.applyTemplate / 直 setSheet 通路验证(不依赖 prop 二次接线)
// 为简化,我们通过命令式 API 操控已加载的工作簿:
//   ① applyTemplate({placeholders:{name:'Alice'},anchors:[...]}) 在样例上叠一道填充,验证文本/锚点单元格
//   ② applyTemplate 配合 ABORT signal:预取消应立即抛 AbortError(API 暴露过来)
function run(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: applyTemplate(占位符+锚点)在样例工作簿上生效`, async ({ page }) => {
    await page.goto(url)
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => (window as any)[h]?.getWorkbook?.() != null, handle, { timeout: 20_000 })

    // 已经加载的样例:A1 = "2026 年度销售汇总"(string)。先把它改成含 {{}} 的占位符(借 setStyle/insert 不易,
    // 直接用 getWorkbook + 改 cell raw,只为测占位符替换路径)
    await page.evaluate((h) => {
      const wb = (window as any)[h].getWorkbook()
      const sheet = wb.sheets[0]
      // cellKey 'r:c'
      const cell = sheet.cells.get('0:0')
      cell.raw = '客户: {{customer}} - 季度: {{quarter}}'
      cell.type = 'string'
    }, handle)

    // applyTemplate:占位符 + 锚点表(往 A12 起填两行)
    const result = await page.evaluate(async (h) => {
      return await (window as any)[h].applyTemplate({
        placeholders: { customer: 'Alice', quarter: 'Q1' },
        anchors: [{ startCell: 'A12', rows: [['apple', 1], ['banana', 2]] }],
      })
    }, handle)
    expect(result.anchorsWritten).toBe(4)

    // 验证占位符替换
    const a1 = await page.evaluate((h) => (window as any)[h].getCellValue(0, 0), handle)
    expect(a1).toBe('客户: Alice - 季度: Q1')

    // 验证锚点写入
    const a12 = await page.evaluate((h) => (window as any)[h].getCellValue(11, 0), handle)
    expect(a12).toBe('apple')
    const b12 = await page.evaluate((h) => (window as any)[h].getCellValue(11, 1), handle)
    expect(b12).toBe(1)
    const a13 = await page.evaluate((h) => (window as any)[h].getCellValue(12, 0), handle)
    expect(a13).toBe('banana')
  })

  test(`${label}: applyTemplate 配合 pre-aborted signal 抛 AbortError`, async ({ page }) => {
    await page.goto(url)
    await page.getByRole('button', { name: '加载示例' }).click()
    await expect(page.locator(canvasSel)).toBeVisible({ timeout: 20_000 })
    await page.waitForFunction((h) => (window as any)[h]?.getWorkbook?.() != null, handle, { timeout: 20_000 })

    const r = await page.evaluate(async (h) => {
      const ctrl = new AbortController()
      ctrl.abort()
      try {
        await (window as any)[h].applyTemplate({
          placeholders: { x: 'y' },
          signal: ctrl.signal,
        })
        return { ok: false, name: null as string | null }
      } catch (e) {
        return { ok: true, name: (e as Error & { name: string }).name }
      }
    }, handle)
    expect(r.ok).toBe(true)
    expect(r.name).toBe('AbortError')
  })
}

test.describe('JSON 直渲 + 模板填值 e2e(P3)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
})
