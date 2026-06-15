/**
 * 自动填充序列引擎(框架无关,纯函数,可单测)。1.10.0 新增。
 *
 * 给定一段"源值"(按填充方向排好序)和要填的格数,产出接续模式的新值。对齐 Excel/WPS 拖拽填充柄:
 *  - 全数值:1 个 → 复制;≥2 个 → 等差外推(步长 = 相邻差均值)
 *  - 全日期:1 个 → 每格 +1 天;≥2 个 → 按相邻差(天)外推
 *  - "前缀+末尾整数"文本(如 "Item 1" / "第1周"):递增末尾整数,保留前缀
 *  - 星期/月份名(中英常见写法):按位置循环接续
 *  - 其它:循环复制源值
 *
 * 方向无关:总是把 source 当作"正方向"序列向后接续。向上/向左填充由调用方反转 source + 结果。
 */
import type { CellValue } from '../model/data-access'

const DAY_MS = 86_400_000

// 循环型名称表(检测到源全在某表里 → 按表循环接续)
const CYCLES: string[][] = [
  ['周一', '周二', '周三', '周四', '周五', '周六', '周日'],
  ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'],
  ['一月', '二月', '三月', '四月', '五月', '六月', '七月', '八月', '九月', '十月', '十一月', '十二月'],
  ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
  ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
]

/**
 * 接续 source 产出 count 个新值。source 已按填充正方向排序(至少 1 个)。
 *
 * `ctrl` = 拖拽时是否按住 Ctrl,翻转"复制 ↔ 序列"(对齐 Excel/WPS):
 *  - 数值:普通 单个→复制 / ≥2→等差;Ctrl 单个→递增(+1) / ≥2→复制
 *  - 日期/星期月份/文本递增:普通→序列;Ctrl→复制
 *  - 纯文本:始终循环复制(Ctrl 无影响)
 */
export function computeFillSeries(source: CellValue[], count: number, ctrl = false): CellValue[] {
  if (count <= 0) return []
  const src = source.length ? source : [null]
  const copy = () => Array.from({ length: count }, (_, i) => src[i % src.length])

  // 1. 全数值
  if (src.every((v) => typeof v === 'number')) {
    const nums = src as number[]
    const wantSeries = (nums.length >= 2) !== ctrl // XOR:普通 ≥2 才序列;Ctrl 翻转
    if (!wantSeries) return copy()
    const step = nums.length >= 2 ? avgStep(nums) : 1 // Ctrl+单个 → 步长 1
    const start = nums[nums.length - 1]
    return Array.from({ length: count }, (_, i) => start + step * (i + 1))
  }

  // 2. 全日期(普通 → 序列;Ctrl → 复制)
  if (src.every((v) => v instanceof Date)) {
    if (ctrl) return copy()
    const ms = (src as Date[]).map((d) => d.getTime())
    const step = ms.length >= 2 ? avgStep(ms) : DAY_MS
    const start = ms[ms.length - 1]
    return Array.from({ length: count }, (_, i) => new Date(start + step * (i + 1)))
  }

  // 3. 循环名称表(星期/月份)(普通 → 接续;Ctrl → 复制)
  const cyc = detectCycle(src)
  if (cyc) {
    if (ctrl) return copy()
    const { list, lastIdx, step } = cyc
    return Array.from({ length: count }, (_, i) => list[mod(lastIdx + step * (i + 1), list.length)])
  }

  // 4. 前缀 + 末尾整数(普通 → 递增;Ctrl → 复制)
  const tn = detectTrailingNumber(src)
  if (tn) {
    if (ctrl) return copy()
    const { prefix, suffix, lastNum, step, pad } = tn
    return Array.from({ length: count }, (_, i) => `${prefix}${padNum(lastNum + step * (i + 1), pad)}${suffix}`)
  }

  // 5. 兜底:循环复制源值
  return copy()
}

// ---------------- 内部 ----------------

function avgStep(nums: number[]): number {
  let sum = 0
  for (let i = 1; i < nums.length; i++) sum += nums[i] - nums[i - 1]
  return sum / (nums.length - 1)
}

function mod(n: number, m: number): number {
  return ((n % m) + m) % m
}

function detectCycle(src: CellValue[]): { list: string[]; lastIdx: number; step: number } | null {
  if (!src.every((v) => typeof v === 'string')) return null
  const strs = src as string[]
  for (const list of CYCLES) {
    const idxs = strs.map((s) => list.indexOf(s))
    if (idxs.some((i) => i < 0)) continue
    const step = idxs.length >= 2 ? Math.round(avgStep(idxs.map((i) => unwrap(i, idxs[0], list.length)))) || 1 : 1
    return { list, lastIdx: idxs[idxs.length - 1], step }
  }
  return null
}
// 把循环索引去回绕(如 [6,0] 视作 [6,7] 求步长 +1)
function unwrap(i: number, first: number, len: number): number {
  return i < first ? i + len : i
}

function detectTrailingNumber(src: CellValue[]): { prefix: string; suffix: string; lastNum: number; step: number; pad: number } | null {
  if (!src.every((v) => typeof v === 'string')) return null
  const strs = src as string[]
  const re = /^(.*?)(\d+)(\D*)$/
  const parsed = strs.map((s) => re.exec(s))
  if (parsed.some((m) => !m)) return null
  const ms = parsed as RegExpExecArray[]
  const prefix = ms[ms.length - 1][1]
  const suffix = ms[ms.length - 1][3]
  // 前后缀需一致才算同一序列(否则当复制)
  if (!ms.every((m) => m[1] === prefix && m[3] === suffix)) return null
  const nums = ms.map((m) => parseInt(m[2], 10))
  const step = nums.length >= 2 ? Math.round(avgStep(nums)) || 1 : 1
  const lastDigits = ms[ms.length - 1][2]
  const pad = lastDigits.length > 1 && lastDigits[0] === '0' ? lastDigits.length : 0 // 保留前导零位宽
  return { prefix, suffix, lastNum: nums[nums.length - 1], step, pad }
}

function padNum(n: number, pad: number): string {
  const s = String(Math.abs(n))
  const body = pad > 0 ? s.padStart(pad, '0') : s
  return n < 0 ? '-' + body : body
}
