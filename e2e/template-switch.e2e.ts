import { test, expect, type Page } from '@playwright/test'

/**
 * 模板样式 overlay e2e(P3 重设计 2026-06-08)
 *
 * 新语义:
 *   - JSON 数据源 (`:workbook`) + 模板 → JSON 在 A1 自然位置;模板贡献 styling
 *     (字体/边框/列宽/合并/freeze);模板的装饰文字 (订单结算单/客户/合计/{{占位符}}) **全部丢弃**
 *   - xlsx 数据源 (`:src`) + 模板 → 模板**忽略** (xlsx 自带格式)
 */

async function loadJsonSample(page: Page) {
  await page.goto('/')
  await page.getByRole('button', { name: 'JSON 示例' }).click()
  await page.waitForFunction(() => (window as any).__excelViewer?.getWorkbook?.() != null, undefined, { timeout: 20_000 })
}

test.describe('JSON 数据源 / 模板样式 overlay(P3 重设计)', () => {
  test('JSON 示例 (无模板) → 数据在 A1 自然位置;模板状态为默认', async ({ page }) => {
    await loadJsonSample(page)

    // jsonToWorkbook 对象数组 → 自动 headerRow:A1:E1 表头 (keys) + A2 起数据
    const cells = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      return {
        a1: v.getCellText(0, 0), // 'name' (表头)
        b1: v.getCellText(0, 1), // 'price'
        a2: v.getCellText(1, 0), // '笔记本电脑' (第一行数据)
        a3: v.getCellText(2, 0), // '机械键盘'
        a5existsAsRowFive: v.getCellText(4, 0), // 第五行数据 = 第 4 个 item '鼠标'
      }
    })
    expect(cells.a1).toBe('name')
    expect(cells.b1).toBe('price')
    expect(cells.a2).toBe('笔记本电脑')
    expect(cells.a3).toBe('机械键盘')
    expect(cells.a5existsAsRowFive).toBe('鼠标')
  })

  test('JSON 示例 → 导入模板 → 模板装饰文字全部不见, JSON 仍在 A1 自然位置', async ({ page }) => {
    await loadJsonSample(page)

    // 标题区 fileName 显示自定义名
    const titleBefore = await page.locator('.toolbar .file').textContent()
    expect(titleBefore).toContain('订单数据')
    expect(titleBefore).not.toContain('模板:')

    // 工具栏「模板 ▾」拾取 public/template-sample.xlsx
    const fileInput = page.locator('input[type=file][accept=".xlsx,.xlsm"][hidden]').nth(1)
    await fileInput.setInputFiles('public/template-sample.xlsx')

    // 等渲染应用 (workbook 已被 applyStyleTemplate 合并过) —— 通过 templateName 后缀触发
    await page.waitForFunction(() => {
      const t = document.querySelector('.toolbar .file')?.textContent ?? ''
      return t.includes('模板:')
    }, undefined, { timeout: 20_000 })

    const titleAfter = await page.locator('.toolbar .file').textContent()
    expect(titleAfter).toContain('订单数据')
    expect(titleAfter).toContain('模板: template-sample.xlsx')

    // ★ 关键回归: 数据仍在 A1 自然位置 (模板没把它"挤"到下面)
    const cells = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      return {
        a1: v.getCellText(0, 0), // 'name' 表头, JSON 在 A1
        a2: v.getCellText(1, 0), // '笔记本电脑'
        a5: v.getCellText(4, 0), // '鼠标'
      }
    })
    expect(cells.a1).toBe('name')
    expect(cells.a2).toBe('笔记本电脑')
    expect(cells.a5).toBe('鼠标')

    // ★ 模板装饰文字全部丢弃 —— 把整张表所有 raw 收集起来, 都不应该包含模板的装饰文字
    const allRaws = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      const wb = v.getWorkbook()
      const sheet = wb.sheets[0]
      return [...sheet.cells.values()].map((c: any) => c.raw)
    })
    expect(allRaws).not.toContain('订单结算单')
    expect(allRaws).not.toContain('客户:')
    expect(allRaws).not.toContain('{{customer}}')
    expect(allRaws).not.toContain('{{total}}')
    expect(allRaws).not.toContain('合计:')

    // 注: 模板的具体 styling 指纹 (列宽 / 合并 / 样式池) 在 src/core/template/__tests__/style-overlay.test.ts
    // 已经覆盖, e2e 这里只验证用户能感知的行为 (数据位置 + 装饰文字丢弃 + 标题后缀)

    // 切回默认渲染: 工具栏「模板 ▾」→「清除模板」
    const tplBtn = page.locator('.action-toolbar .dd .tool', { hasText: '模板' })
    await tplBtn.click()
    await page.locator('.tb-menu').getByText('清除模板').click()

    await page.waitForFunction(() => {
      const t = document.querySelector('.toolbar .file')?.textContent ?? ''
      return !t.includes('模板:')
    }, undefined, { timeout: 20_000 })

    const titleCleared = await page.locator('.toolbar .file').textContent()
    expect(titleCleared).toContain('订单数据')
    expect(titleCleared).not.toContain('模板:')

    // 清除后, 数据仍然在 A1 (从未被挤过)
    const a1After = await page.evaluate(() => (window as any).__excelViewer.getCellText(0, 0))
    expect(a1After).toBe('name')
  })

  test('xlsx 数据源加载示例后, 工具栏「模板 ▾」应被禁用', async ({ page }) => {
    await page.goto('/')
    await page.getByRole('button', { name: '加载示例' }).click()
    await page.waitForFunction(() => (window as any).__excelViewer?.getWorkbook?.() != null, undefined, { timeout: 20_000 })

    // 模板按钮存在但 disabled (xlsx 数据源不支持套模板) —— .dd 包裹的才是可见行,measure 行不算
    const tplBtn = page.locator('.action-toolbar .dd .tool', { hasText: '模板' })
    await expect(tplBtn).toBeDisabled()
  })
})
