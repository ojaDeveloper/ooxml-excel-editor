import { test, expect, type Page } from '@playwright/test'

/**
 * editableTargets 白名单 API e2e (2026-06-08).
 *
 * 流程:
 *   ① 加载示例 → 开编辑模式 → 默认所有格 isCellEditable=true
 *   ② 命令式 setEditableTargets([{row:0,col:0},{row:2,col:2}]) → 只这两格可编辑, 其它 false
 *   ③ setEditableTargets(undefined) → 关闭白名单 → 恢复全可编辑
 *   ④ Demo「设置可编辑」按钮 → 弹窗点选 → 应用 → 同样的白名单生效
 */

async function loadSampleEditable(page: Page) {
  await page.goto('/')
  await page.getByRole('button', { name: '加载示例' }).click()
  await page.waitForFunction(() => (window as any).__excelViewer?.getWorkbook?.() != null, undefined, { timeout: 20_000 })
  // 开编辑模式 (label 的可见文本是 "编辑模式")
  await page.getByText('编辑模式').click()
}

test.describe('editableTargets 白名单 e2e (2026-06-08)', () => {
  test('命令式 setEditableTargets: 4 种 target 形状 + 关闭白名单', async ({ page }) => {
    await loadSampleEditable(page)

    // 默认: 整表可编辑
    const before = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      return {
        a1: v.isCellEditable(0, 0),
        c3: v.isCellEditable(2, 2),
        e5: v.isCellEditable(4, 4),
        whitelist: v.getEditableTargets(),
      }
    })
    expect(before.a1).toBe(true)
    expect(before.c3).toBe(true)
    expect(before.e5).toBe(true)
    expect(before.whitelist).toBeUndefined()

    // 设白名单: 不相邻 2 格 + 1 整行 + 1 矩形
    await page.evaluate(() => {
      ;(window as any).__excelViewer.setEditableTargets([
        { row: 0, col: 0 }, // 单格 A1
        { row: 2, col: 5 }, // 单格 F3
        { row: 4 },          // 整行 R5
        { top: 6, left: 1, bottom: 7, right: 3 }, // 矩形 B7:D8
      ])
    })

    const after = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      return {
        a1: v.isCellEditable(0, 0),
        a2: v.isCellEditable(1, 0),
        f3: v.isCellEditable(2, 5),
        g3: v.isCellEditable(2, 6),
        r5c0: v.isCellEditable(4, 0),
        r5c99: v.isCellEditable(4, 99),
        b7: v.isCellEditable(6, 1),
        d8: v.isCellEditable(7, 3),
        e8: v.isCellEditable(7, 4),
        getCount: Array.isArray(v.getEditableTargets()) ? (v.getEditableTargets() as unknown[]).length : -1,
      }
    })
    expect(after.a1).toBe(true)   // 单格命中
    expect(after.a2).toBe(false)  // 单格外
    expect(after.f3).toBe(true)   // 单格命中
    expect(after.g3).toBe(false)  // 单格外
    expect(after.r5c0).toBe(true) // 整行任一列命中
    expect(after.r5c99).toBe(true) // 整行任一列命中
    expect(after.b7).toBe(true)   // 矩形命中
    expect(after.d8).toBe(true)   // 矩形右下角 (闭区间)
    expect(after.e8).toBe(false)  // 矩形外
    expect(after.getCount).toBe(4)

    // 关闭白名单 → 恢复整表可编辑
    await page.evaluate(() => (window as any).__excelViewer.setEditableTargets(undefined))
    const cleared = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      return {
        a1: v.isCellEditable(0, 0),
        e5: v.isCellEditable(4, 4),
        whitelist: v.getEditableTargets(),
      }
    })
    expect(cleared.a1).toBe(true)
    expect(cleared.e5).toBe(true)
    expect(cleared.whitelist).toBeUndefined()
  })

  test('Demo 「设置可编辑」按钮: 弹窗点选 → 应用 → 白名单生效', async ({ page }) => {
    await loadSampleEditable(page)

    // 点工具栏「设置可编辑」(在 demo bar, 编辑模式下出现)
    await page.getByRole('button', { name: '设置可编辑', exact: true }).click()
    // 弹窗已开
    await expect(page.locator('.edit-targets-dialog')).toBeVisible()

    // 点 R1C1 (第 1 行第 1 列) + R3C3 (第 3 行第 3 列)
    const r1c1 = page.locator('.edit-targets-grid tbody tr').nth(0).locator('td').nth(0)
    const r3c3 = page.locator('.edit-targets-grid tbody tr').nth(2).locator('td').nth(2)
    await r1c1.click()
    await r3c3.click()

    // 点列标题 E → 整列 E (col=4) 可编辑
    await page.locator('.edit-targets-grid thead th', { hasText: 'E' }).click()

    // 应用
    await page.getByRole('button', { name: '应用', exact: true }).click()
    await expect(page.locator('.edit-targets-dialog')).toBeHidden()

    const got = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      return {
        a1: v.isCellEditable(0, 0),   // 勾选了
        b1: v.isCellEditable(0, 1),   // 没勾
        c3: v.isCellEditable(2, 2),   // 勾选了 (R3C3)
        c4: v.isCellEditable(3, 2),   // 没勾
        e1: v.isCellEditable(0, 4),   // 整列 E 命中
        e99: v.isCellEditable(99, 4), // 整列 E 命中
      }
    })
    expect(got.a1).toBe(true)
    expect(got.b1).toBe(false)
    expect(got.c3).toBe(true)
    expect(got.c4).toBe(false)
    expect(got.e1).toBe(true)
    expect(got.e99).toBe(true)
  })

  test('显式空数组 [] → 全只读 (没格在白名单)', async ({ page }) => {
    await loadSampleEditable(page)
    await page.evaluate(() => (window as any).__excelViewer.setEditableTargets([]))
    const got = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      return { a1: v.isCellEditable(0, 0), e5: v.isCellEditable(4, 4) }
    })
    expect(got.a1).toBe(false)
    expect(got.e5).toBe(false)
  })
})
