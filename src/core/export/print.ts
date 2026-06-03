/**
 * 打印: 把每个工作表的合成底图分页后塞进一个临时打印窗口,调用 window.print()。
 * 用户可在系统打印对话框里"另存为 PDF",故零依赖即可覆盖 PDF 需求。
 * headerHtml / footerHtml 出现在每个打印页(扩展点)。
 */
import type { PrintOptions } from './types'
import type { ExportSheetImage } from './pdf'
import { resolveMargins, resolvePageSize } from './raster'
import { sliceToPages } from './paginate'

export function printSheets(sheets: ExportSheetImage[], opts: PrintOptions = {}): void {
  if (!sheets.length) return
  const [pageW, pageH] = resolvePageSize(opts.format, opts.orientation)
  const margin = resolveMargins(opts.margin)
  const contentWmm = Math.max(1, pageW - margin.left - margin.right)
  const contentHmm = Math.max(1, pageH - margin.top - margin.bottom)
  const fitToWidth = opts.fitToWidth ?? true
  const title = opts.title || '打印'

  const pagesHtml: string[] = []
  for (const sh of sheets) {
    const slices = sliceToPages(sh.canvas, sh.bodyWcss, sh.bodyHcss, {
      contentWmm,
      contentHmm,
      fitToWidth,
      zoom: sh.zoom,
      repeatTop: sh.repeatTop,
    })
    for (const pg of slices) {
      pagesHtml.push(
        `<div class="page">` +
          (opts.headerHtml ? `<div class="hdr">${opts.headerHtml}</div>` : '') +
          `<img src="${pg.canvas.toDataURL('image/png')}" style="width:${pg.widthMm}mm;height:${pg.heightMm}mm" />` +
          (opts.footerHtml ? `<div class="ftr">${opts.footerHtml}</div>` : '') +
          `</div>`,
      )
    }
  }
  if (!pagesHtml.length) return

  const html =
    `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>` +
    `@page{size:${pageW}mm ${pageH}mm;margin:${margin.top}mm ${margin.right}mm ${margin.bottom}mm ${margin.left}mm}` +
    `*{box-sizing:border-box}body{margin:0;font-family:Calibri,Arial,sans-serif}` +
    `.page{page-break-after:always;break-after:page;overflow:hidden}` +
    `.page:last-child{page-break-after:auto;break-after:auto}` +
    `img{display:block}.hdr{font-size:11px;color:#555;margin-bottom:4px}.ftr{font-size:11px;color:#555;margin-top:4px}` +
    `</style></head><body>${pagesHtml.join('')}</body></html>`

  const w = window.open('', '_blank')
  if (!w) {
    console.warn('[ooxml-preview] 打印窗口被浏览器拦截,请允许弹窗')
    return
  }
  w.document.open()
  w.document.write(html)
  w.document.close()
  // 图片以 dataURL 内联,通常已就绪;延时一拍确保布局完成再打印
  const fire = () => {
    w.focus()
    w.print()
  }
  if (w.document.readyState === 'complete') setTimeout(fire, 200)
  else w.onload = () => setTimeout(fire, 200)
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}
