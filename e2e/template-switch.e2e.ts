import { test, expect, type Page } from '@playwright/test'

async function ready(page: Page, canvasSel: string, handle: string) {
  // 点击「JSON 示例」按钮(Vue)/「加载示例」(React 仅 .xlsx)
  await expect(page.locator(canvasSel)).toBeAttached({ timeout: 20_000 })
}

const call = (page: Page, handle: string, fn: string, ...args: unknown[]) =>
  page.evaluate(([h, f, a]) => (window as any)[h][f as string](...(a as unknown[])), [handle, fn, args] as const)

// P3 进阶:displayFileName + templateName 后缀 + 运行时模板切换
test.describe('JSON 数据源 / 模板切换 e2e(P3 进阶)', () => {
  test(`Vue: JSON 示例 → 标题显示自定义名;工具栏「模板 ▾」拾取 → fileName 后缀加 · 模板: xxx`, async ({ page }) => {
    await page.goto('/')
    // 点 JSON 示例(Vue demo 才有这个按钮)
    await page.getByRole('button', { name: 'JSON 示例' }).click()
    await page.waitForFunction(() => (window as any).__excelViewer?.getWorkbook?.() != null, undefined, { timeout: 20_000 })

    // 标题:JSON 数据源时 :fileName="'订单结算单'" → 这个名显示
    const titleVue = await page.locator('.toolbar .file').textContent()
    expect(titleVue).toContain('订单结算单')
    expect(titleVue).not.toContain('模板:')

    // 工具栏「模板 ▾」拾取 public/template-sample.xlsx → 后缀加 · 模板: xxx
    // 直接通过隐藏 input 上传(模拟 click 后的 file picker)
    const fileInput = page.locator('input[type=file][accept=".xlsx,.xlsm"][hidden]').nth(1) // 第 1 个是 .file-btn,第 2 个是模板
    await fileInput.setInputFiles('public/template-sample.xlsx')
    // 等模板加载完
    await page.waitForFunction(
      () => {
        const wb = (window as any).__excelViewer?.getWorkbook?.()
        return wb && wb.sheets[0]?.name === '发票模板' // 模板的表名
      },
      undefined,
      { timeout: 20_000 },
    )

    const titleVueAfter = await page.locator('.toolbar .file').textContent()
    expect(titleVueAfter).toContain('订单结算单')
    expect(titleVueAfter).toContain('模板: template-sample.xlsx')

    // 切回默认渲染:从模板下拉里点「清除模板」
    // 模板按钮在工具栏里;它是一个下拉,先点 caret 展开
    // 工具栏「模板」下拉(避开隐藏 measure 行;取可见的、有 active 类的)
    const tplBtn = page.locator('.action-toolbar .dd .tool', { hasText: '模板' })
    await tplBtn.click()
    await page.locator('.tb-menu').getByText('清除模板').click()

    // 等渲染回 JSON
    await page.waitForFunction(
      () => {
        const wb = (window as any).__excelViewer?.getWorkbook?.()
        return wb && wb.sheets[0]?.name !== '发票模板'
      },
      undefined,
      { timeout: 20_000 },
    )
    const titleVueCleared = await page.locator('.toolbar .file').textContent()
    expect(titleVueCleared).toContain('订单结算单')
    expect(titleVueCleared).not.toContain('模板:')
  })
})
