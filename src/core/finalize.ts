/** 主线程后处理: 图片字节 → blob URL;文件格式探测;错误信息友好化。 */
import type { WorkbookModel } from './model/types'

/** 把解析层产出的图片字节转成 blob URL(只能在有 URL.createObjectURL 的主线程做) */
export function finalizeImages(model: WorkbookModel): void {
  if (typeof URL === 'undefined' || !URL.createObjectURL) return
  for (const sheet of model.sheets) {
    for (const img of sheet.images) {
      if (img.bytes && img.mime && !img.src) {
        img.src = URL.createObjectURL(new Blob([img.bytes as BlobPart], { type: img.mime }))
      }
    }
  }
  // WPS 单元格内嵌图(DISPIMG)登记表同样落 blob url
  if (model.cellImages) {
    for (const ci of model.cellImages.values()) {
      if (ci.bytes && ci.mime && !ci.src) {
        ci.src = URL.createObjectURL(new Blob([ci.bytes as BlobPart], { type: ci.mime }))
      }
    }
  }
}

/** 释放之前 finalizeImages 生成的 blob URL,避免内存泄漏 */
export function revokeImages(model: WorkbookModel): void {
  if (typeof URL === 'undefined' || !URL.revokeObjectURL) return
  for (const sheet of model.sheets) {
    for (const img of sheet.images) {
      if (img.src?.startsWith('blob:')) URL.revokeObjectURL(img.src)
    }
  }
  if (model.cellImages) {
    for (const ci of model.cellImages.values()) {
      if (ci.src?.startsWith('blob:')) URL.revokeObjectURL(ci.src)
    }
  }
}

export type FormatKind = 'xlsx' | 'xls' | 'not-zip' | 'empty'

/** 看文件头几个字节判断格式,在真正解析前就能给出友好提示 */
export function detectFormat(buffer: ArrayBuffer): FormatKind {
  const b = new Uint8Array(buffer)
  if (b.length < 4) return 'empty'
  // OLE2/CFB 复合文档头 → 旧版 .xls / 加密 ooxml
  if (b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0) return 'xls'
  // ZIP 头 "PK" → xlsx/xlsm
  if (b[0] === 0x50 && b[1] === 0x4b) return 'xlsx'
  return 'not-zip'
}

/** 把底层异常翻成用户能懂的话 */
export function friendlyError(e: unknown): string {
  const msg = (e as any)?.message || String(e)
  if (/encrypt|password|protected/i.test(msg)) return '文件已加密或受密码保护，暂不支持预览。'
  if (/zip|signature|central directory|invalid|corrupt|unexpected end/i.test(msg))
    return '文件无法解析：可能不是有效的 .xlsx，或文件已损坏。'
  return '解析失败：' + msg
}
