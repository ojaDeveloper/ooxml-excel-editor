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

// Excel/WPS 把格式放 <style> 类里(<td class="xl65"> + .xl65{...}),不是内联 style=
const HTML_CLASS =
  '<html><head><style>' +
  '.xl65{font-weight:700;background:#FFFF00;color:#FF0000}' +
  '.xl66{border-top:.5pt solid windowtext;text-align:center}' +
  '</style></head><body><table>' +
  '<tr><td class="xl65" x:num="9">9</td></tr>' +
  '<tr><td class="xl66">y</td></tr>' +
  '</table></body></html>'

function runClass(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 富粘贴 Excel/WPS 类样式(<style> .xl 类)→ 边框/填充/字体/对齐还原`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    expect(await call(page, handle, 'pasteRichHtml', HTML_CLASS, { row: 2, col: 1 })).toBe(true)
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBe(9) // x:num 原始值

    // .xl65 类:粗体 + 黄填充 + 红字(若不解析 <style> 类则全丢)
    const s1 = (await call(page, handle, 'getCellSnapshot', 2, 1)) as any
    expect(s1.style.font.bold).toBe(true)
    expect(String(s1.style.fill.fgColor).toUpperCase()).toBe('#FFFF00')
    expect(String(s1.style.font.color).toUpperCase()).toBe('#FF0000')
    // .xl66 类:上边框 + 居中
    const s2 = (await call(page, handle, 'getCellSnapshot', 3, 1)) as any
    expect(s2.style.borders.top?.style).toBe('thin')
    expect(s2.style.hAlign).toBe('center')
  })
}

// VML o:gfxdata 夹具:zip 里含 1x1 PNG(scripts 生成)
const VML_GFX =
  'UEsDBBQAAAAIAK2Uy1wuc8SQOwAAAEMAAAAaAAAAY2xpcGJvYXJkL21lZGlhL2ltYWdlMS5wbmfrDPBz5+WS4mJgYOD19HAJAtKMIMzBBiTlRY90AikuTxfHkIo5ySAZVgZGXi7dLUBRBk9XP5d1TglNAFBLAQIUABQAAAAIAK2Uy1wuc8SQOwAAAEMAAAAaAAAAAAAAAAAAAAAAAAAAAABjbGlwYm9hcmQvbWVkaWEvaW1hZ2UxLnBuZ1BLBQYAAAAAAQABAEgAAABzAAAAAAA='

// 真实 WPS 区域复制:格式在 <style> 类(含 mso-number-format),图片在 VML o:gfxdata(file:/// 的 <img> 读不了)
const HTML_WPS =
  '<html><head><meta name=Generator content="Microsoft Excel"><style>\n<!--td\n\t{vertical-align:middle;white-space:nowrap;font-size:11.0pt;mso-number-format:General;}\n.et4\n\t{color:#0D0D0D;mso-number-format:"yyyy/m/d";border:.5pt solid #000000;background:#FFFFFF;text-align:center;white-space:normal;}\n.et8\n\t{border:.5pt solid #000000;}\n-->\n</style></head><body><table>' +
  "<col width=72 style='width:54.00pt;'><col width=120 style='width:90.00pt;'>" + // 列宽
  "<tr height=88 style='height:66.00pt;'>" + // 行高
  '<td class=et4 width=72 x:num="46113" style=\'width:54.00pt;\'>2026/4/1</td>' +
  '<td class=et8 width=120 style=\'width:90.00pt;\'>' +
  `<!--[if gte vml 1]><v:shape id="ID_X" type="#_x0000_t75" o:gfxdata="${VML_GFX}"></v:shape><![endif]-->` +
  '<![if !vml]><span><img width=69 height=70 src="file:///C:/temp/clip.png"></span><![endif]>' +
  '</td></tr></table></body></html>'

function runWps(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: WPS 区域复制 — 类数字格式(日期序列号→日期)+ VML o:gfxdata 内嵌图还原`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    expect(await call(page, handle, 'pasteRichHtml', HTML_WPS, { row: 2, col: 1 })).toBe(true)

    // 日期格:x:num 序列号 46113 + 类 mso-number-format "yyyy/m/d" → 值是数字 46113,但带上日期格式码
    expect(await call(page, handle, 'getCellValue', 2, 1)).toBe(46113)
    const s1 = (await call(page, handle, 'getCellSnapshot', 2, 1)) as any
    expect(s1.style.numFmt).toBe('yyyy/m/d') // 不再丢成裸序列号
    expect(s1.style.borders.top?.style).toBe('thin')
    expect(s1.style.hAlign).toBe('center') // text-align:center 还原
    expect(s1.style.wrapText).toBe(true) // white-space:normal → 自动换行还原(否则长文本溢出,连居中也看不出)
    expect(s1.style.vAlign).toBe('middle') // 裸 td{vertical-align:middle} 默认层 → 垂直居中还原(et4 自己没写)
    expect(s1.style.font.size).toBe(11) // td{font-size:11.0pt} 按 pt 解析(不被当 px 算成 8)

    // VML o:gfxdata 内嵌图 → 解 zip 取 png → 落成单元格图(workbook 登记表里有一张)
    // 行高搬到目标行;列宽**故意不搬**(整列共享,改了会动表头)→ 列宽保持粘贴前不变
    const m = await page.evaluate((h) => {
      const v = (window as any)[h]
      const wb = v.getWorkbook()
      const sh = wb.sheets[v.getActiveSheet()]
      return { cellImages: wb.cellImages ? wb.cellImages.size : 0, disp: !!sh.cells.get('2:2')?.dispImgId, col1: sh.columns.get(1)?.width, col2: sh.columns.get(2)?.width, row2: sh.rows.get(2)?.height }
    }, handle)
    expect(m.cellImages).toBeGreaterThan(0)
    expect(m.disp).toBe(true)
    // 行高 1:1 搬到目标行(<tr height=88>)
    expect(m.row2).toBe(88)
    // 列宽没有被 WPS 的 <col width=72/120> 覆盖(否则会等于 72/120 而非现有表头宽度)
    expect(m.col1).not.toBe(72)
    expect(m.col2).not.toBe(120)

    // 例外:粘到**首行**(row 0,上方无表头可破坏)→ 套用源列宽(<col width=72/120>),粘贴块成新表头/新布局
    expect(await call(page, handle, 'pasteRichHtml', HTML_WPS, { row: 0, col: 1 })).toBe(true)
    const top = await page.evaluate((h) => {
      const v = (window as any)[h]
      const sh = v.getWorkbook().sheets[v.getActiveSheet()]
      return { col1: sh.columns.get(1)?.width, col2: sh.columns.get(2)?.width }
    }, handle)
    expect(top.col1).toBe(72)
    expect(top.col2).toBe(120)
  })
}

// 真实 Ctrl+V 路径:派发 paste 事件(带原始 HTML)到 scroller → onPaste → 拿原始 HTML(不像 clipboard.read 那样净化)
function runPasteEvent(label: string, url: string, canvasSel: string, handle: string, scrollerSel: string) {
  test(`${label}: Ctrl+V(paste 事件,原始 HTML)→ WPS 类格式 + numFmt + VML 图都还原(不被净化)`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    const applied = await page.evaluate(
      ({ h, html, sel }) => {
        const v = (window as any)[h]
        v.setSelection({ top: 2, left: 1, bottom: 2, right: 1 })
        const scroller = document.querySelector(sel) as HTMLElement
        const dt = new DataTransfer()
        dt.setData('text/html', html)
        scroller.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }))
        const s = v.getCellSnapshot(2, 1)
        const wb = v.getWorkbook()
        return { value: v.getCellValue(2, 1), numFmt: s?.style?.numFmt, borderTop: s?.style?.borders?.top?.style, cellImages: wb.cellImages ? wb.cellImages.size : 0 }
      },
      { h: handle, html: HTML_WPS, sel: scrollerSel },
    )
    expect(applied.value).toBe(46113)
    expect(applied.numFmt).toBe('yyyy/m/d') // 数字格式没被净化删掉
    expect(applied.borderTop).toBe('thin') // 类样式没被净化删掉
    expect(applied.cellImages).toBeGreaterThan(0) // VML 内嵌图没被净化删掉
  })
}

// 覆盖式样式:粘到**已有底色**的多行区上,源里没写填充的格不应保留目标底色(干净覆盖,贴近源,同 Excel)
// 3 行 × 2 列:每行都有一个"显式白底"格 + 一个"只边框没写填充"格 —— 验证**每一行**都覆盖,不止首行
const HTML_OVERRIDE =
  '<table>' +
  '<tr><td style="background:#FFFFFF">a1</td><td style="border:1px solid #000000">b1</td></tr>' +
  '<tr><td style="background:#FFFFFF">a2</td><td style="border:1px solid #000000">b2</td></tr>' +
  '<tr><td style="background:#FFFFFF">a3</td><td style="border:1px solid #000000">b3</td></tr>' +
  '</table>'

function runOverride(label: string, url: string, canvasSel: string, handle: string) {
  test(`${label}: 粘到带底色的多行区 → 每一行都覆盖式(源没写填充的格清掉目标底色,不漏色)`, async ({ page }) => {
    await page.goto(url)
    await ready(page, canvasSel, handle)

    // 先给目标 (5..7, 5..6) 整片套红底
    await call(page, handle, 'setStyle', { top: 5, left: 5, bottom: 7, right: 6 }, { fill: { type: 'solid', fgColor: '#FF0000' } })
    for (const r of [5, 6, 7]) expect(String(((await call(page, handle, 'getCellSnapshot', r, 6)) as any).style.fill.fgColor).toUpperCase()).toBe('#FF0000')

    // 粘 3 行到 (5,5):每行 a=显式白底 → 白;b=只边框没写填充 → 应清成无填充(不保留红底)
    expect(await call(page, handle, 'pasteRichHtml', HTML_OVERRIDE, { row: 5, col: 5 })).toBe(true)
    for (let i = 0; i < 3; i++) {
      const r = 5 + i
      const a = (await call(page, handle, 'getCellSnapshot', r, 5)) as any
      const b = (await call(page, handle, 'getCellSnapshot', r, 6)) as any
      expect(String(a.style.fill.fgColor).toUpperCase()).toBe('#FFFFFF') // 源白底覆盖红底(第 1/2/3 行都是)
      expect(b.style.fill.type).toBe('none') // 源没写填充 → 干净覆盖成无填充,不漏目标红底(第 1/2/3 行都是)
      expect(b.style.borders.top?.style).toBe('thin') // 边框照常还原
    }
  })
}

test.describe('富粘贴 e2e(Phase C:Excel/WPS HTML → 样式/合并 + 单次撤销)', () => {
  run('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  run('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
  runClass('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  runClass('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
  runWps('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  runWps('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
  runOverride('Vue', '/', 'canvas.grid-canvas', '__excelViewer')
  runOverride('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact')
  runPasteEvent('Vue', '/', 'canvas.grid-canvas', '__excelViewer', '.scroller')
  runPasteEvent('React', '/react.html', 'canvas.rxl-canvas', '__excelViewerReact', '.rxl-scroller')
})
