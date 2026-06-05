/**
 * 图片放大灯箱(框架无关 DOM)—— 挂到 document.body 的全屏暗背景 + 居中大图 + 下载/关闭。
 * 点背景 / Esc / 关闭按钮关闭(点图本身不关)。Vue / React 壳只在点击图片时让 controller 调 show,
 * 灯箱逻辑全在这。仿 ContextMenuHost 的 body 级浮层范式。
 */
import { downloadBlob } from '../export/raster'

export interface LightboxImage {
  /** 图片地址(blob url / data url / http);直接喂 <img src> */
  src: string
  /** 下载时的文件名(缺省按 mime 推扩展名) */
  fileName?: string
  mime?: string
}

export class LightboxHost {
  private el: HTMLElement | null = null
  private cleanup: (() => void) | null = null

  isOpen(): boolean {
    return this.el !== null
  }

  show(img: LightboxImage): void {
    if (typeof document === 'undefined') return
    this.close()
    const backdrop = document.createElement('div')
    backdrop.className = 'ooxml-lightbox'
    backdrop.style.cssText =
      'position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.72);' +
      'display:flex;align-items:center;justify-content:center;'

    const image = document.createElement('img')
    image.src = img.src
    image.draggable = false
    image.style.cssText =
      'max-width:88vw;max-height:82vh;object-fit:contain;background:#fff;box-shadow:0 8px 40px rgba(0,0,0,.5);'
    backdrop.appendChild(image)

    const bar = document.createElement('div')
    bar.style.cssText = 'position:fixed;top:16px;right:20px;display:flex;gap:10px;'
    const dl = makeBtn('⬇ 下载原图')
    dl.addEventListener('mousedown', (e) => {
      e.preventDefault()
      e.stopPropagation()
      void this.download(img)
    })
    const cl = makeBtn('✕ 关闭')
    cl.addEventListener('mousedown', (e) => {
      e.preventDefault()
      e.stopPropagation()
      this.close()
    })
    bar.append(dl, cl)
    backdrop.appendChild(bar)

    // 点背景空白处关闭(点图/工具条不关)
    backdrop.addEventListener('mousedown', (e) => {
      if (e.target === backdrop) this.close()
    })
    document.body.appendChild(backdrop)
    this.el = backdrop

    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') this.close()
    }
    document.addEventListener('keydown', onKey, true)
    this.cleanup = () => document.removeEventListener('keydown', onKey, true)
  }

  private async download(img: LightboxImage): Promise<void> {
    try {
      const blob = await (await fetch(img.src)).blob()
      const ext = (img.mime?.split('/')[1] || 'png').replace('jpeg', 'jpg').replace('svg+xml', 'svg')
      downloadBlob(blob, img.fileName || `image.${ext}`)
    } catch {
      /* 下载失败(跨域 blob 等)忽略 */
    }
  }

  close(): void {
    this.cleanup?.()
    this.cleanup = null
    this.el?.remove()
    this.el = null
  }
  dispose(): void {
    this.close()
  }
}

function makeBtn(text: string): HTMLButtonElement {
  const b = document.createElement('button')
  b.textContent = text
  b.style.cssText =
    'padding:6px 12px;border:none;border-radius:6px;background:rgba(255,255,255,.92);color:#1f2329;' +
    "font:13px/1.4 -apple-system,'Segoe UI',sans-serif;cursor:pointer;"
  return b
}
