import { test, expect, type Page } from '@playwright/test'
import { loadSample } from './helpers'

// 真系统剪贴板往返(需读写权限)。验证本组件自己复制 → 粘贴走 data-ooxml-clip 快照做 1:1。
test.use({ permissions: ['clipboard-read', 'clipboard-write'] })

test.describe('复制粘贴 1:1 保真(走剪贴板嵌入快照)', () => {
  test('Ctrl+C 复制数字格 → Ctrl+V 粘贴:仍是数字(不退化成文本)且值一致', async ({ page }) => {
    await loadSample(page)
    await page.waitForFunction(() => (window as any).__excelViewer?.rectOf?.(1, 0) != null, null, { timeout: 5_000 })
    await page.getByText('编辑模式').click()

    const area = page.locator('.render-area')
    const box = (await area.boundingBox())!
    const clickCell = async (r: number, c: number) => {
      const rect = await page.evaluate(([rr, cc]) => (window as any).__excelViewer.rectOf(rr, cc), [r, c] as const)
      await page.mouse.click(box.x + rect.x + rect.w / 2, box.y + rect.y + rect.h / 2)
    }

    // 源:单价列 (2,1) 是数字
    const srcVal = await page.evaluate(() => (window as any).__excelViewer.getCellValue(2, 1))
    expect(typeof srcVal).toBe('number')

    await clickCell(2, 1) // 选中并聚焦
    await page.keyboard.press('Control+C')
    await page.waitForTimeout(200) // 等异步剪贴板写入完成

    await clickCell(12, 6) // 空白目标
    await page.keyboard.press('Control+V')
    await page.waitForFunction(() => (window as any).__excelViewer.getCellValue(12, 6) != null, null, { timeout: 5_000 })

    const pasted = await page.evaluate(() => {
      const v = (window as any).__excelViewer
      const val = v.getCellValue(12, 6)
      return { val, type: typeof val }
    })
    expect(pasted.type).toBe('number') // 关键:没变成格式化文本
    expect(pasted.val).toBe(srcVal)
  })
})
