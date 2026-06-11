/**
 * 右键上下文菜单宿主(框架无关 DOM)—— 挂到 document.body 的 fixed 菜单,点项执行 action + 关闭,
 * 点外部 / Esc 关闭,贴边自动翻转。支持一级子菜单(`children`,悬停展开右侧 flyout)。
 * Vue / React 壳只把 contextmenu 事件转给 controller,菜单逻辑全在这。
 */
export interface MenuItem {
  label?: string
  action?: () => void
  disabled?: boolean
  /** true = 分隔线 */
  separator?: boolean
  /** 子菜单项(有则该行不执行 action,悬停展开右侧 flyout) */
  children?: MenuItem[]
}

const MENU_CSS =
  'position:fixed;z-index:9999;background:#fff;border:1px solid #d0d3d7;border-radius:6px;' +
  'box-shadow:0 4px 16px rgba(0,0,0,.16);padding:4px 0;min-width:168px;' +
  "font:13px/1.6 -apple-system,'Segoe UI',sans-serif;user-select:none;color:#1f2329;"

export class ContextMenuHost {
  private el: HTMLElement | null = null
  private submenus: HTMLElement[] = [] // 当前打开的子菜单(flyout),关闭时一并清掉
  private openSubRow: HTMLElement | null = null // 当前 flyout 所属的父行(避免重复打开)
  private subCloseTimer: ReturnType<typeof setTimeout> | null = null // 延时关闭(给鼠标从父行移到 flyout 的缓冲)
  private cleanup: (() => void) | null = null

  isOpen(): boolean {
    return this.el !== null
  }

  show(x: number, y: number, items: MenuItem[]): void {
    this.close()
    const menu = this.buildMenu(items)
    document.body.appendChild(menu)
    this.el = menu

    // 贴边翻转:超出视口右/下边则向左/上对齐
    const w = menu.offsetWidth
    const h = menu.offsetHeight
    const left = x + w > window.innerWidth ? Math.max(0, x - w) : x
    const top = y + h > window.innerHeight ? Math.max(0, y - h) : y
    menu.style.left = left + 'px'
    menu.style.top = top + 'px'

    const onDocDown = (ev: MouseEvent) => {
      const t = ev.target as Node
      if (this.el && !this.el.contains(t) && !this.submenus.some((s) => s.contains(t))) this.close()
    }
    const onKey = (ev: KeyboardEvent) => {
      if (ev.key === 'Escape') this.close()
    }
    const onScroll = () => this.close()
    // 延后注册,避免触发它的这次 mousedown 立刻关掉自己
    setTimeout(() => {
      document.addEventListener('mousedown', onDocDown, true)
      document.addEventListener('keydown', onKey, true)
      window.addEventListener('scroll', onScroll, true)
      this.cleanup = () => {
        document.removeEventListener('mousedown', onDocDown, true)
        document.removeEventListener('keydown', onKey, true)
        window.removeEventListener('scroll', onScroll, true)
      }
    }, 0)
  }

  /** 建一级菜单 DOM(不挂载、不定位);子菜单悬停时再建 flyout。 */
  private buildMenu(items: MenuItem[]): HTMLElement {
    const menu = document.createElement('div')
    menu.className = 'ooxml-context-menu'
    menu.style.cssText = MENU_CSS
    for (const it of items) {
      if (it.separator) {
        const hr = document.createElement('div')
        hr.style.cssText = 'height:1px;background:#eceef1;margin:4px 0;'
        menu.appendChild(hr)
        continue
      }
      const row = document.createElement('div')
      const hasKids = !!it.children?.length
      row.textContent = (it.label ?? '') + (hasKids ? '  ▸' : '')
      const dis = !!it.disabled
      row.style.cssText = `padding:5px 16px;cursor:${dis ? 'default' : 'pointer'};color:${dis ? '#b7bcc2' : 'inherit'};white-space:nowrap;${hasKids ? 'display:flex;justify-content:space-between;gap:18px;' : ''}`
      if (!dis) {
        row.addEventListener('mouseenter', () => {
          row.style.background = '#eef3fe'
          this.cancelSubClose()
          if (hasKids) {
            if (this.openSubRow !== row) { this.closeAllSubmenus(); this.openSubmenu(row, it.children as MenuItem[]) }
          } else {
            // 进了没有子菜单的兄弟行 → 延时关掉已开的 flyout(延时给"斜着滑进 flyout"留缓冲)
            if (this.openSubRow) this.scheduleSubClose()
          }
        })
        row.addEventListener('mouseleave', () => {
          row.style.background = ''
          if (hasKids && this.openSubRow === row) this.scheduleSubClose() // 离开父行 → 延时关(移进 flyout 会取消)
        })
        if (!hasKids) {
          row.addEventListener('mousedown', (e) => {
            e.preventDefault()
            e.stopPropagation()
            this.close()
            it.action?.()
          })
        }
      }
      menu.appendChild(row)
    }
    return menu
  }

  /** 在 row 右侧展开子菜单 flyout(贴边翻转到左侧);悬停 flyout 取消关闭,离开 flyout 延时关。 */
  private openSubmenu(row: HTMLElement, children: MenuItem[]): void {
    const sub = this.buildMenu(children)
    sub.style.visibility = 'hidden'
    document.body.appendChild(sub)
    const r = row.getBoundingClientRect()
    const w = sub.offsetWidth
    const h = sub.offsetHeight
    // 紧贴父菜单右缘(-2 抵消两边边框,避免可见缝隙致鼠标滑过时丢失 hover);溢出则翻到左侧
    const left = r.right + w > window.innerWidth ? Math.max(0, r.left - w + 2) : r.right - 2
    const top = r.top + h > window.innerHeight ? Math.max(0, window.innerHeight - h) : r.top
    sub.style.left = left + 'px'
    sub.style.top = top + 'px'
    sub.style.visibility = 'visible'
    sub.addEventListener('mouseenter', () => this.cancelSubClose())
    sub.addEventListener('mouseleave', () => this.scheduleSubClose())
    this.submenus.push(sub)
    this.openSubRow = row
  }

  private scheduleSubClose(): void {
    this.cancelSubClose()
    this.subCloseTimer = setTimeout(() => this.closeAllSubmenus(), 260)
  }
  private cancelSubClose(): void {
    if (this.subCloseTimer) { clearTimeout(this.subCloseTimer); this.subCloseTimer = null }
  }
  private closeAllSubmenus(): void {
    this.cancelSubClose()
    for (const s of this.submenus) s.remove()
    this.submenus = []
    this.openSubRow = null
  }

  close(): void {
    this.cleanup?.()
    this.cleanup = null
    this.closeAllSubmenus()
    this.el?.remove()
    this.el = null
  }
  dispose(): void {
    this.close()
  }
}
