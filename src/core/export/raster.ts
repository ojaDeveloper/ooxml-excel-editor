/** 位图工具: canvas ↔ blob / dataURL,下载,图片加载,纸张尺寸换算。 */
import type { Margins, Orientation, PageFormat } from './types'

const MIME: Record<string, string> = { png: 'image/png', jpeg: 'image/jpeg', webp: 'image/webp' }

export function canvasToBlob(canvas: HTMLCanvasElement, type = 'png', quality = 0.92): Promise<Blob> {
  const mime = MIME[type] ?? 'image/png'
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => (blob ? resolve(blob) : reject(new Error('canvas.toBlob 返回空'))),
      mime,
      mime === 'image/png' ? undefined : quality,
    )
  })
}

export function canvasToDataURL(canvas: HTMLCanvasElement, type = 'png', quality = 0.92): string {
  const mime = MIME[type] ?? 'image/png'
  return canvas.toDataURL(mime, mime === 'image/png' ? undefined : quality)
}

/** 触发浏览器下载一个 Blob */
export function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  document.body.appendChild(a)
  a.click()
  a.remove()
  // 给浏览器一点时间发起下载再回收
  setTimeout(() => URL.revokeObjectURL(url), 4000)
}

/** 从 URL / blobURL / dataURL 加载为 HTMLImageElement(用于合成到导出底图) */
export function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.crossOrigin = 'anonymous'
    img.onload = () => resolve(img)
    img.onerror = () => reject(new Error('图片加载失败: ' + src))
    img.src = src
  })
}

// ---------------- 纸张尺寸(mm) ----------------
const FORMATS: Record<string, [number, number]> = {
  a4: [210, 297],
  a3: [297, 420],
  letter: [215.9, 279.4],
}

/** 解析纸张为 [宽, 高] mm,应用方向 */
export function resolvePageSize(format: PageFormat = 'a4', orientation: Orientation = 'portrait'): [number, number] {
  const base = Array.isArray(format) ? format : FORMATS[format] ?? FORMATS.a4
  const [w, h] = base
  return orientation === 'landscape' ? [Math.max(w, h), Math.min(w, h)] : [Math.min(w, h), Math.max(w, h)]
}

/** 归一化页边距(数字 → 四边;部分 → 补默认 10) */
export function resolveMargins(margin: number | Partial<Margins> | undefined): Margins {
  if (typeof margin === 'number') return { top: margin, right: margin, bottom: margin, left: margin }
  return { top: margin?.top ?? 10, right: margin?.right ?? 10, bottom: margin?.bottom ?? 10, left: margin?.left ?? 10 }
}

/** css px → mm(按 96dpi) */
export const MM_PER_PX = 25.4 / 96
