import { test, expect, type Page } from '@playwright/test'
import { loadSample } from './helpers'

/**
 * 透视表 e2e:工具栏入口 → 创建对话框 → 静态结果 + 右侧字段面板;
 * createPivotTable API(新建工作表);exportXlsx 回注真实 OOXML pivot 零件。
 * 示例文件「销售报表」数据区 A2:E7(表头行 1,数据行 2..6),列:产品/单价/数量/金额/增长率。
 */
const SRC = { top: 1, left: 0, bottom: 6, right: 4 } as const

async function loadEditable(page: Page) {
  await loadSample(page)
  await page.waitForFunction(() => (window as any).__excelViewer?.rectOf?.(1, 0) != null, null, { timeout: 5_000 })
  await page.getByText('编辑模式').click()
}

test.describe('透视表(WPS 式入口 + 字段面板 + API + OOXML 导出)', () => {
  test('工具栏入口:对话框创建空白透视表 → 面板勾选字段填充 → 切换汇总 → undo 撤销', async ({ page }) => {
    await loadEditable(page)
    await page.evaluate((sel) => (window as any).__excelViewer.setSelection(sel), SRC)

    // 入口按钮 → 创建对话框(默认输出 = 选区右侧空两列 G2)
    await page.locator('.action-toolbar').getByRole('button', { name: /透视表/ }).click()
    await expect(page.locator('[data-cell]')).toHaveValue('G2')
    await page.locator('[data-ok]').click()

    // 空白起步(对齐 WPS):G2 = (1,6) 只是占位框,不猜字段
    expect(await page.evaluate(() => (window as any).__excelViewer.getCellText(1, 6))).toBe('数据透视表')

    // 右侧字段面板出现;勾选字段复选框:产品(col0,文本→行)、金额(col3,数值→值)
    const panel = page.locator('div', { hasText: '数据透视表区域' }).last()
    await expect(panel).toBeVisible()
    await page.locator('[data-toggle-field="0"]').check()
    await page.locator('[data-toggle-field="3"]').check()
    const header = await page.evaluate(() => [(window as any).__excelViewer.getCellText(1, 6), (window as any).__excelViewer.getCellText(1, 7)])
    expect(header[0]).toBe('产品')
    expect(header[1]).toContain('求和项')

    // 值字段切换 求和 → 计数 → 结果重建为行数计数
    await page.locator('[data-summary-value]').selectOption('count')
    const counted = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      let totalRow = -1
      for (let r = 1; r < 14; r++) if (v.getCellText(r, 6) === '总计') { totalRow = r; break }
      return { header: v.getCellText(1, 7), total: v.getCellValue(totalRow, 7) }
    })
    expect(counted.header).toContain('计数项')
    expect(counted.total).toBe(5) // 5 行数据 → 计数总计 5

    // undo → 整个创建撤销(回到创建前快照,面板字段改动属派生态,一并消失)
    const afterUndo = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      v.undo()
      return v.getCellText(1, 6)
    })
    expect(afterUndo).toBe('')
  })

  test('createPivotTable API:新建工作表输出 + 金额求和正确;导出 xlsx 含真实 OOXML pivot 零件', async ({ page }) => {
    await loadEditable(page)

    const r = await page.evaluate((sel) => {
      const v = (window as any).__excelViewer
      const expected = [2, 3, 4, 5, 6].reduce((s, row) => s + Number(v.getCellValue(row, 3) ?? 0), 0)
      const ok = v.createPivotTable({
        sourceRange: sel,
        output: { kind: 'new-sheet' },
        layout: { rows: [0], values: [{ field: 3, summary: 'sum' }] },
      })
      // 新表激活后读静态结果(表头 + 总计行)
      const texts: string[] = []
      for (let row = 0; row < 10; row++) texts.push(v.getCellText(row, 0))
      const totalRow = texts.findIndex((t) => t === '总计')
      return { ok, header: v.getCellText(0, 0), valueHeader: v.getCellText(0, 1), total: v.getCellValue(totalRow, 1), expected }
    }, SRC)
    expect(r.ok).toBe(true)
    expect(r.header).toBe('产品')
    expect(r.valueHeader).toContain('金额')
    expect(r.total).toBe(r.expected)

    // 导出 .xlsx → zip 里有标准 pivot 零件(zip 头里文件名是明文,直接搜字节)
    const names = await page.evaluate(async () => {
      const blob: Blob = await (window as any).__excelViewer.exportXlsx()
      const bytes = new Uint8Array(await blob.arrayBuffer())
      const text = new TextDecoder('latin1').decode(bytes)
      return {
        table: text.includes('xl/pivotTables/pivotTable1.xml'),
        cache: text.includes('xl/pivotCache/pivotCacheDefinition1.xml'),
        records: text.includes('xl/pivotCache/pivotCacheRecords1.xml'),
      }
    })
    expect(names.table).toBe(true)
    expect(names.cache).toBe(true)
    expect(names.records).toBe(true)

    // 新建工作表后,顶部 sheet-tabs 应出现并高亮 PivotTable 标签(回归:shallowRef 不通知导致新表不显示)
    const tabs = page.locator('.sheet-tabs .tab')
    await expect(tabs.filter({ hasText: 'PivotTable' })).toHaveClass(/active/)
  })

  // 读 H 列里 '总计' 行在 I 列(col 8)的值
  const grandTotal = (page: Page) => page.evaluate(() => {
    const v = (window as any).__excelViewer
    for (let r = 1; r < 60; r++) if (v.getCellText(r, 7) === '总计') return v.getCellValue(r, 8) as number
    return null
  })

  test('活刷新:编辑源数据格 → 透视结果自动重算', async ({ page }) => {
    await loadEditable(page)
    const ok = await page.evaluate((sel) => (window as any).__excelViewer.createPivotTable({
      sourceRange: sel,
      output: { kind: 'current-sheet', cell: 'H2' },
      layout: { rows: [0], values: [{ field: 3, summary: 'sum' }] },
    }), SRC)
    expect(ok).toBe(true)
    const before = await grandTotal(page)
    expect(before).not.toBeNull()

    // 源数据「金额」列(col 3)首数据行(row 2)+1000 → 透视总计应同步 +1000
    const orig = await page.evaluate(() => Number((window as any).__excelViewer.getCellValue(2, 3) ?? 0))
    await page.evaluate((v) => (window as any).__excelViewer.editCell(2, 3, v + 1000), orig)
    expect(await grandTotal(page)).toBe((before as number) + 1000)

    // 撤销源编辑 → 透视也跟着还原
    await page.evaluate(() => (window as any).__excelViewer.undo())
    expect(await grandTotal(page)).toBe(before)
  })

  test('2 行字段:折叠按钮隐藏明细,再点展开', async ({ page }) => {
    await loadEditable(page)
    // 产品(0) 外层 + 数量(2) 内层 → 可折叠的大纲
    const ok = await page.evaluate((sel) => (window as any).__excelViewer.createPivotTable({
      sourceRange: sel,
      output: { kind: 'current-sheet', cell: 'H2' },
      layout: { rows: [0, 2], values: [{ field: 3, summary: 'sum' }] },
    }), SRC)
    expect(ok).toBe(true)
    const countRows = () => page.evaluate(() => {
      const v = (window as any).__excelViewer
      let n = 0
      for (let r = 1; r < 60; r++) { n++; if (v.getCellText(r, 7) === '总计') break }
      return n
    })
    const before = await countRows()

    // 点第一个分组表头(H3 = row 2)左侧的折叠按钮
    const area = page.locator('.render-area')
    const abox = (await area.boundingBox())!
    const rect = await page.evaluate(() => (window as any).__excelViewer.rectOf(2, 7))
    await page.mouse.click(abox.x + rect.x + 7, abox.y + rect.y + rect.h / 2)
    const collapsed = await countRows()
    expect(collapsed).toBeLessThan(before) // 明细被折叠隐藏

    // 再点一次 → 展开还原
    await page.mouse.click(abox.x + rect.x + 7, abox.y + rect.y + rect.h / 2)
    expect(await countRows()).toBe(before)
  })

  test('多选(include)筛选:只保留选中的产品参与汇总', async ({ page }) => {
    await loadEditable(page)
    // 头两个产品名(数据行 2、3 的 A 列)
    const products = await page.evaluate(() => [2, 3].map((r) => (window as any).__excelViewer.getCellText(r, 0)))
    const expected = await page.evaluate(() => [2, 3].reduce((s, r) => s + Number((window as any).__excelViewer.getCellValue(r, 3) ?? 0), 0))

    const ok = await page.evaluate(([sel, prods]) => (window as any).__excelViewer.createPivotTable({
      sourceRange: sel,
      output: { kind: 'current-sheet', cell: 'H2' },
      layout: { rows: [0], values: [{ field: 3, summary: 'sum' }], filters: [{ field: 0, mode: 'include', values: prods }] },
    }), [SRC, products] as const)
    expect(ok).toBe(true)

    // 总计 = 仅选中两个产品的金额之和;行分组数 = 2(+ 表头 + 总计 = 4 行)
    expect(await grandTotal(page)).toBe(expected)
    const rowCount = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      let n = 0
      for (let r = 1; r < 60; r++) { n++; if (v.getCellText(r, 7) === '总计') break }
      return n
    })
    expect(rowCount).toBe(4)
  })

  test('列字段:横向展开成二维交叉表(列加进来后表变宽)', async ({ page }) => {
    await loadEditable(page)
    // 计算 H2 起表头行(row 1, col 7 起)的非空宽度
    const headerWidth = () => page.evaluate(() => {
      const v = (window as any).__excelViewer
      let w = 0
      for (let c = 7; c < 40; c++) { if (v.getCellText(1, c) === '') break; w++ }
      return w
    })
    // 仅行+值:宽度 = 2(行标签 + 1 个值列)
    await page.evaluate((sel) => (window as any).__excelViewer.createPivotTable({
      sourceRange: sel, output: { kind: 'current-sheet', cell: 'H2' },
      layout: { rows: [0], values: [{ field: 3, summary: 'sum' }] },
    }), SRC)
    const noCol = await headerWidth()
    expect(noCol).toBe(2)
    await page.evaluate(() => (window as any).__excelViewer.undo())

    // 加列字段「数量」(col 2)→ 每个不同数量拆一列,表头变宽
    await page.evaluate((sel) => (window as any).__excelViewer.createPivotTable({
      sourceRange: sel, output: { kind: 'current-sheet', cell: 'H2' },
      layout: { rows: [0], columns: [2], values: [{ field: 3, summary: 'sum' }] },
    }), SRC)
    expect(await headerWidth()).toBeGreaterThan(noCol)
  })

  test('字段面板:筛选明细面板渲染多选复选框', async ({ page }) => {
    await loadEditable(page)
    await page.evaluate((sel) => (window as any).__excelViewer.setSelection(sel), SRC)
    await page.locator('.action-toolbar').getByRole('button', { name: /透视表/ }).click()
    await page.locator('[data-ok]').click()
    // 把产品(field 0)加入筛选器 → chip 出现编辑按钮
    await page.locator('[data-add="filters"][data-field="0"]').click()
    await page.locator('[data-filter-edit="0"]').click()
    // 切到「多选」→ 出现值复选框
    await page.locator('[data-filter-mode="include"]').check()
    await expect(page.locator('[data-filter-check]').first()).toBeVisible()
    expect(await page.locator('[data-filter-check]').count()).toBeGreaterThan(1)
  })
})
