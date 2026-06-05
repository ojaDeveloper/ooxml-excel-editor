/**
 * 薄层: 用 fflate 解压 .xlsx(zip) 取出原始 XML part，用 fast-xml-parser 转 JS 对象。
 * 专门用来读 ExcelJS 丢弃/不完整的部分: theme 主题色、drawings 锚点、charts。
 */
import { unzipSync, strFromU8 } from 'fflate'
import { XMLParser } from 'fast-xml-parser'

export interface RawPackage {
  files: Record<string, Uint8Array>
  parse(path: string): any | undefined
  text(path: string): string | undefined
  bytes(path: string): Uint8Array | undefined
  list(prefix: string): string[]
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true, // 去掉命名空间前缀(a: / c: / xdr:)，统一节点名
  allowBooleanAttributes: true,
  parseAttributeValue: true,
  trimValues: true,
  // 放宽实体展开上限:大表(几百个 DISPIMG 格,每个含 &quot;)会超 fxp 默认 1000 上限而抛错,
  // 导致 drawing-parser / row-meta 在大表上静默失败。本地解析可信 xlsx,放宽到足够大。
  processEntities: {
    enabled: true,
    maxEntitySize: 10_000_000,
    maxTotalExpansions: 100_000_000,
    maxExpandedLength: 100_000_000,
    maxEntityCount: 10_000_000,
  },
})

export function openPackage(buffer: ArrayBuffer): RawPackage {
  const files = unzipSync(new Uint8Array(buffer))
  const cache = new Map<string, any>()
  const norm = (p: string) => p.replace(/^\//, '')
  return {
    files,
    bytes(path) {
      return files[norm(path)]
    },
    text(path) {
      const b = files[norm(path)]
      return b ? strFromU8(b) : undefined
    },
    parse(path) {
      const key = norm(path)
      if (cache.has(key)) return cache.get(key)
      const b = files[key]
      if (!b) return undefined
      const obj = parser.parse(strFromU8(b))
      cache.set(key, obj)
      return obj
    },
    list(prefix) {
      const p = norm(prefix)
      return Object.keys(files).filter((f) => f.startsWith(p))
    },
  }
}

/** 解析 part 的 .rels，返回 r:id → target(相对该 part 目录解析后的绝对包内路径) */
export function parseRels(pkg: RawPackage, partPath: string): Record<string, string> {
  const dir = partPath.includes('/') ? partPath.slice(0, partPath.lastIndexOf('/')) : ''
  const relsPath = dir ? `${dir}/_rels/${basename(partPath)}.rels` : `_rels/${basename(partPath)}.rels`
  const xml = pkg.parse(relsPath)
  const out: Record<string, string> = {}
  if (!xml?.Relationships?.Relationship) return out
  const rels = toArray(xml.Relationships.Relationship)
  for (const r of rels) {
    const id = r['@_Id']
    let target = r['@_Target'] as string
    if (!id || !target) continue
    if (r['@_TargetMode'] === 'External') {
      out[id] = target // 外部链接保留原样
    } else {
      out[id] = resolvePath(dir, target)
    }
  }
  return out
}

export function basename(p: string): string {
  return p.slice(p.lastIndexOf('/') + 1)
}

/** 相对 baseDir 解析 target(支持 ../) → 包内绝对路径(无前导 /) */
export function resolvePath(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const stack = baseDir ? baseDir.split('/') : []
  for (const seg of target.split('/')) {
    if (seg === '..') stack.pop()
    else if (seg === '.' || seg === '') continue
    else stack.push(seg)
  }
  return stack.join('/')
}

/** fast-xml-parser 对单元素/多元素返回 object/array 不一致，统一成数组 */
export function toArray<T = any>(v: T | T[] | undefined | null): T[] {
  if (v === undefined || v === null) return []
  return Array.isArray(v) ? v : [v]
}
