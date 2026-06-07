import { describe, it, expect } from 'vitest'
import { checkAborted, isAbortError, yieldToEvent } from '../abort'

describe('checkAborted', () => {
  it('signal 未取消时不抛', () => {
    const c = new AbortController()
    expect(() => checkAborted(c.signal)).not.toThrow()
  })

  it('signal.abort() 后抛 AbortError(name 字段稳定,便于上层 catch 分流)', () => {
    const c = new AbortController()
    c.abort()
    try {
      checkAborted(c.signal)
      throw new Error('checkAborted 应当抛错')
    } catch (e) {
      expect(isAbortError(e)).toBe(true)
    }
  })

  it('未传 signal 时不抛', () => {
    expect(() => checkAborted(undefined)).not.toThrow()
  })
})

describe('yieldToEvent', () => {
  it('返回 Promise<number>;node 环境走 setTimeout fallback,不挂死', async () => {
    const t = await yieldToEvent()
    expect(typeof t).toBe('number')
  })
})

describe('isAbortError', () => {
  it('普通 Error 不算', () => {
    expect(isAbortError(new Error('x'))).toBe(false)
  })
  it('name=AbortError 算', () => {
    const e: Error & { name: string } = new Error('x') as any
    e.name = 'AbortError'
    expect(isAbortError(e)).toBe(true)
  })
})
