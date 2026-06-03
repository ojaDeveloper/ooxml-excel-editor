/**
 * PDF 导出: 把每个工作表的合成底图竖向分页贴进 PDF。
 * jspdf 为"可选 peer 依赖",动态 import;未安装时抛友好错误(其余导出能力不受影响)。
 * 每页贴图后调用 beforeRenderPage(ctx) —— 宿主/插件可在此画页眉/页脚/水印/页码(扩展点)。
 */
import type { PdfExportOptions, PdfPageContext } from './types'
import { resolveMargins, resolvePageSize } from './raster'
import { sliceToPages } from './paginate'

/** 一个待导出工作表的合成底图 */
export interface ExportSheetImage {
  /** 已合成(格子 + 图片/图表/形状)的底图,设备像素 */
  canvas: HTMLCanvasElement
  /** 正文 css 宽高(zoom=1,不含表头) */
  bodyWcss: number
  bodyHcss: number
  sheetName: string
}

async function loadJsPdf(): Promise<any> {
  try {
    // jspdf 为可选 peer 依赖,可能未安装 → 运行时动态加载,类型与打包都不强依赖
    // @ts-ignore 可选依赖,宿主未装时由 catch 给出友好提示
    const mod: any = await import(/* @vite-ignore */ 'jspdf')
    return mod.jsPDF || mod.default?.jsPDF || mod.default || mod
  } catch (e) {
    throw new Error('PDF 导出需要可选依赖 jspdf,请先安装: npm i jspdf (' + (e as Error).message + ')')
  }
}

/** 生成 PDF,返回 Blob(不触发下载)。 */
export async function exportToPdf(sheets: ExportSheetImage[], opts: PdfExportOptions = {}): Promise<Blob> {
  if (!sheets.length) throw new Error('没有可导出的工作表')
  const JsPDF = await loadJsPdf()
  const [pageW, pageH] = resolvePageSize(opts.format, opts.orientation)
  const margin = resolveMargins(opts.margin)
  const contentWmm = Math.max(1, pageW - margin.left - margin.right)
  const contentHmm = Math.max(1, pageH - margin.top - margin.bottom)
  const fitToWidth = opts.fitToWidth ?? true

  // 先把每个表切片,算总页数(给钩子用)
  const perSheet = sheets.map((sh) => ({
    sheet: sh,
    pages: sliceToPages(sh.canvas, sh.bodyWcss, sh.bodyHcss, { contentWmm, contentHmm, fitToWidth }),
  }))
  const pageCount = perSheet.reduce((n, s) => n + Math.max(s.pages.length, 0), 0) || 1

  const doc = new JsPDF({ unit: 'mm', format: [pageW, pageH], orientation: 'portrait' })
  let globalPage = 0
  let first = true

  for (let si = 0; si < perSheet.length; si++) {
    const { sheet, pages } = perSheet[si]
    if (!pages.length) continue
    for (const pg of pages) {
      if (!first) doc.addPage([pageW, pageH], 'portrait')
      first = false
      const dataUrl = pg.canvas.toDataURL('image/png')
      doc.addImage(dataUrl, 'PNG', margin.left, margin.top, pg.widthMm, pg.heightMm, undefined, 'FAST')
      if (opts.beforeRenderPage) {
        const ctx: PdfPageContext = {
          doc,
          pageIndex: globalPage,
          pageCount,
          sheetName: sheet.sheetName,
          sheetIndex: si,
          pageWidth: pageW,
          pageHeight: pageH,
          margin,
        }
        try {
          opts.beforeRenderPage(ctx)
        } catch (e) {
          console.warn('[ooxml-preview] beforeRenderPage 抛错:', e)
        }
      }
      globalPage++
    }
  }
  return doc.output('blob')
}
