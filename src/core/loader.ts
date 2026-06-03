/** 把多种输入归一化成 ArrayBuffer,供解析层使用。可选 onProgress 报字节读取进度。 */

export type ExcelSource = File | Blob | ArrayBuffer | Uint8Array | string

/** 读取进度回调: loaded/total 字节(total 为 0 表示未知) */
export type ReadProgressFn = (loaded: number, total: number) => void

export async function loadArrayBuffer(src: ExcelSource, onProgress?: ReadProgressFn): Promise<ArrayBuffer> {
  if (src instanceof ArrayBuffer) {
    onProgress?.(src.byteLength, src.byteLength)
    return src
  }
  if (src instanceof Uint8Array) {
    onProgress?.(src.byteLength, src.byteLength)
    return src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength) as ArrayBuffer
  }
  if (typeof Blob !== 'undefined' && src instanceof Blob) {
    return await readBlob(src, onProgress)
  }
  if (typeof src === 'string') {
    return await fetchWithProgress(src, onProgress)
  }
  throw new Error('不支持的输入类型，期望 File/Blob/ArrayBuffer/Uint8Array/URL 字符串')
}

/** Blob/File: 用 FileReader 拿真实读取进度 */
function readBlob(blob: Blob, onProgress?: ReadProgressFn): Promise<ArrayBuffer> {
  if (!onProgress || typeof FileReader === 'undefined') return blob.arrayBuffer()
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onprogress = (e) => {
      if (e.lengthComputable) onProgress(e.loaded, e.total)
    }
    reader.onload = () => {
      onProgress(blob.size, blob.size)
      resolve(reader.result as ArrayBuffer)
    }
    reader.onerror = () => reject(reader.error ?? new Error('文件读取失败'))
    reader.readAsArrayBuffer(blob)
  })
}

/** URL: 流式读取 + Content-Length 算进度 */
async function fetchWithProgress(url: string, onProgress?: ReadProgressFn): Promise<ArrayBuffer> {
  const resp = await fetch(url)
  if (!resp.ok) throw new Error(`加载文件失败: ${resp.status} ${resp.statusText}`)
  const total = Number(resp.headers.get('content-length') || 0)
  if (!onProgress || !resp.body) return await resp.arrayBuffer()

  const reader = resp.body.getReader()
  const chunks: Uint8Array[] = []
  let loaded = 0
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
    loaded += value.byteLength
    onProgress(loaded, total)
  }
  // 合并
  const out = new Uint8Array(loaded)
  let off = 0
  for (const c of chunks) {
    out.set(c, off)
    off += c.byteLength
  }
  onProgress(loaded, total || loaded)
  return out.buffer
}
