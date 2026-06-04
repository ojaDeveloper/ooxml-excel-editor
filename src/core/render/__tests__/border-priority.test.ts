import { describe, it, expect } from 'vitest'
import { heavierEdge } from '../borders'
import type { BorderEdge } from '../../model/types'

const e = (style: BorderEdge['style'], color = '#000000'): BorderEdge => ({ style, color })

/** 相邻单元格共享边:取较重的线型(对齐 Excel/WPS 的边框优先级)。 */
describe('heavierEdge(共享边取较重)', () => {
  it('一侧有、一侧无 → 取有的', () => {
    expect(heavierEdge(e('thin'), undefined)).toEqual(e('thin'))
    expect(heavierEdge(undefined, e('thin'))).toEqual(e('thin'))
    expect(heavierEdge(undefined, undefined)).toBeUndefined()
  })

  it('thick 压 thin、double 压 thick、medium 压 dashed', () => {
    expect(heavierEdge(e('thin'), e('thick'))).toEqual(e('thick'))
    expect(heavierEdge(e('thick'), e('double'))).toEqual(e('double'))
    expect(heavierEdge(e('dashed'), e('medium'))).toEqual(e('medium'))
  })

  it('hair 最轻、被任何实线压过', () => {
    expect(heavierEdge(e('hair'), e('thin'))).toEqual(e('thin'))
    expect(heavierEdge(e('hair'), e('dotted'))).toEqual(e('dotted'))
  })

  it('同线型 → 取本格(a),保留其颜色', () => {
    expect(heavierEdge(e('thin', '#ff0000'), e('thin', '#00ff00'))).toEqual(e('thin', '#ff0000'))
  })

  it('none 视同无', () => {
    expect(heavierEdge(e('none'), e('thin'))).toEqual(e('thin'))
    expect(heavierEdge(e('thin'), e('none'))).toEqual(e('thin'))
  })
})
