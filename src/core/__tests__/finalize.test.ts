import { describe, it, expect } from 'vitest'
import { detectFormat, friendlyError } from '../finalize'

function buf(bytes: number[]): ArrayBuffer {
  return new Uint8Array(bytes).buffer
}

describe('detectFormat', () => {
  it('ZIP 头(PK) → xlsx', () => {
    expect(detectFormat(buf([0x50, 0x4b, 0x03, 0x04]))).toBe('xlsx')
  })
  it('OLE2/CFB 头 → xls(旧格式/加密)', () => {
    expect(detectFormat(buf([0xd0, 0xcf, 0x11, 0xe0]))).toBe('xls')
  })
  it('其它字节 → not-zip', () => {
    expect(detectFormat(buf([0x25, 0x50, 0x44, 0x46]))).toBe('not-zip') // %PDF
  })
  it('过短 → empty', () => {
    expect(detectFormat(buf([0x50]))).toBe('empty')
  })
})

describe('friendlyError', () => {
  it('加密类', () => {
    expect(friendlyError(new Error('file is password protected'))).toContain('加密')
  })
  it('损坏/非 zip 类', () => {
    expect(friendlyError(new Error('invalid zip data / central directory'))).toContain('损坏')
  })
  it('其它原样附带', () => {
    expect(friendlyError(new Error('某个未知错误'))).toBe('解析失败：某个未知错误')
  })
})
