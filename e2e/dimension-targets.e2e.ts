import { test, expect, type Page } from '@playwright/test'

/**
 * Phase B 多形态尺寸 API e2e (2026-06-08)
 *
 * 验证: setColumnWidth / setRowHeight / autoFitColumns / autoFitRows / resetColumnWidth /
 * resetRowHeight 都接 `number | number[] | {from,to}` (DimTarget) 形态; 多 index 单次 undo;
 * strictDimensions=true 时受 editableTargets 白名单约束.
 */

async function loadSampleEditable(page: Page) {
  await page.goto('/')
  await page.getByRole('button', { name: '加载示例' }).click()
  await page.waitForFunction(() => (window as any).__excelViewer?.getWorkbook?.() != null, undefined, { timeout: 20_000 })
  await page.getByText('编辑模式').click()
}

test.describe('Phase B 多形态尺寸 API e2e (2026-06-08)', () => {
  test('setColumnWidth: 单值 / 数组 / 范围 三形态', async ({ page }) => {
    await loadSampleEditable(page)
    const r = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      const got: Record<string, number> = {}
      got.single = v.setColumnWidth(2, 150)         // 单值
      got.array = v.setColumnWidth([4, 6, 8], 88)    // 数组 (不相邻 3 列)
      got.range = v.setColumnWidth({ from: 10, to: 12 }, 60) // 范围 3 列
      // 验证生效
      const wb = v.getWorkbook()
      const sheet = wb.sheets[0]
      got.c3 = sheet.columns.get(2)?.width ?? -1
      got.e5 = sheet.columns.get(4)?.width ?? -1
      got.g7 = sheet.columns.get(6)?.width ?? -1
      got.k11 = sheet.columns.get(10)?.width ?? -1
      got.l12 = sheet.columns.get(11)?.width ?? -1
      got.m13 = sheet.columns.get(12)?.width ?? -1
      return got
    })
    expect(r.single).toBe(1)
    expect(r.array).toBe(3)
    expect(r.range).toBe(3)
    expect(r.c3).toBe(150)
    expect(r.e5).toBe(88)
    expect(r.g7).toBe(88)
    expect(r.k11).toBe(60)
    expect(r.l12).toBe(60)
    expect(r.m13).toBe(60)
  })

  test('setRowHeight 多 index 单次 undo', async ({ page }) => {
    await loadSampleEditable(page)
    const r = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      const sheet = v.getWorkbook().sheets[0]
      const before = [sheet.rows.get(0)?.height, sheet.rows.get(1)?.height, sheet.rows.get(2)?.height]
      v.setRowHeight([0, 1, 2], 40) // 3 行同时改
      const after = [sheet.rows.get(0)?.height, sheet.rows.get(1)?.height, sheet.rows.get(2)?.height]
      v.undo() // 一次撤销应该全恢复
      const restored = [sheet.rows.get(0)?.height, sheet.rows.get(1)?.height, sheet.rows.get(2)?.height]
      return { before, after, restored }
    })
    expect(r.after).toEqual([40, 40, 40])
    // 撤销后 = before
    expect(r.restored[0]).toBe(r.before[0])
    expect(r.restored[1]).toBe(r.before[1])
    expect(r.restored[2]).toBe(r.before[2])
  })

  test('resetColumnWidth: 清除自定义列宽,回落 defaultColWidth', async ({ page }) => {
    await loadSampleEditable(page)
    const r = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      const sheet = v.getWorkbook().sheets[0]
      v.setColumnWidth(3, 200)
      const customW = sheet.columns.get(3)?.width
      const resetCount = v.resetColumnWidth(3)
      const afterReset = sheet.columns.get(3) // 应该 undefined (Map 条目被移除)
      return { customW, resetCount, afterReset }
    })
    expect(r.customW).toBe(200)
    expect(r.resetCount).toBe(1)
    expect(r.afterReset).toBeUndefined()
  })

  test('autoFitColumns: 不传 target → 整表 autoFit', async ({ page }) => {
    await loadSampleEditable(page)
    const result = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      return v.autoFitColumns() // 整表
    })
    expect(typeof result).toBe('number')
    expect(result).toBeGreaterThan(0) // sample 至少有几列
  })

  test('strictDimensions=false (默认): editableTargets 不影响尺寸 API', async ({ page }) => {
    await loadSampleEditable(page)
    const r = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      v.setEditableTargets([{ row: 0, col: 0 }]) // 只 A1 可编辑
      // 默认 strictDimensions=false → 仍可改任意列宽
      return v.setColumnWidth(5, 120)
    })
    expect(r).toBe(1) // 1 列成功
  })

  test('setColumnWidth 负 index 或越界自动跳过', async ({ page }) => {
    await loadSampleEditable(page)
    const r = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      return v.setColumnWidth([-1, -5, 3], 100) // 负值都被 filter 掉, 只剩 3
    })
    expect(r).toBe(1)
  })
})
