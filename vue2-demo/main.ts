/**
 * Vue 2 demo (跟 Vue 3 demo src/App.vue 演示按钮 1:1 对齐).
 * Vue 2 壳内置了 header / action toolbar / 公式栏 / sheet 标签 / 查找; 这里只负责
 * demo 头部的演示按钮 (选择 .xlsx / JSON 示例 / PDF 水印 / 数据→JSON / 贴合 / XLSX/CSV/JSON 导出).
 */
import Vue from 'vue2'
import ExcelViewer from '../src/vue2/ExcelViewer'
import '../src/vue2/excel-viewer.css'
import '../src/demo-shared/demo-bar.css'

new Vue({
  el: '#app',
  components: { ExcelViewer },
  data() {
    return {
      src: undefined as string | File | undefined,
      jsonItems: null as unknown as null | Array<Record<string, unknown>>,
      fileName: '',
      dragOver: false,
      editMode: false,
      highlightReadOnly: false,
      cellImageFit: 'contain' as 'contain' | 'fill' | 'cover',
      lastEvent: '',
      toolbarItems: ['find', 'filter', 'sort', 'clear-filter', 'separator', 'copy', 'pivot-table', 'conditional-format', 'wrap-text', 'image-tools', 'freeze', 'separator', 'template', 'separator', 'zoom', 'export'],
      editableTargetsApplied: undefined as any,
      editTargetsDialogOpen: false,
      editTargetsCells: {} as Record<string, true>,
      editTargetsRows: {} as Record<number, true>,
      editTargetsCols: {} as Record<number, true>,
      // demo 顶栏溢出折叠 (跟 Vue 3 demo 同款 ResizeObserver + measure)
      demoItemWidths: [] as number[],
      demoBarContentW: 0,
      demoMoreOpen: false,
    }
  },
  mounted() {
    this.$nextTick(() => this.demoRemeasure())
    // ResizeObserver / 事件 handler 直接挂 instance, 不入 data → 不被 Vue 转 reactive
    const inst = this as any
    inst._demoRo = new ResizeObserver(() => this.demoRemeasure())
    const bar = this.$refs.demoBarEl as HTMLElement | undefined
    if (bar) inst._demoRo.observe(bar)
    inst._onDocClick = (e: MouseEvent) => {
      const bar2 = this.$refs.demoBarEl as HTMLElement | undefined
      if (bar2 && !bar2.contains(e.target as Node)) this.demoMoreOpen = false
    }
    document.addEventListener('mousedown', inst._onDocClick)
  },
  beforeDestroy() {
    const inst = this as any
    inst._demoRo?.disconnect()
    if (inst._onDocClick) document.removeEventListener('mousedown', inst._onDocClick)
  },
  updated() {
    // DEV: 把命令式 API 挂 window 供 e2e 取几何/读数据(与 Vue3 __excelViewer / React __excelViewerReact 对齐)
    if ((import.meta as any).env?.DEV) (window as any).__excelViewerVue2 = (this.$refs.viewer as any) || null
  },
  watch: {
    // 内容/编辑态变 → 重新测量 (用 watch 而非 updated, 避免 demoRemeasure 改 reactive 触发的死循环)
    editMode() { this.$nextTick(() => this.demoRemeasure()) },
    src() { this.$nextTick(() => this.demoRemeasure()) },
    jsonItems() { this.$nextTick(() => this.demoRemeasure()) },
    highlightReadOnly() { this.$nextTick(() => this.demoRemeasure()) },
    appliedCount() { this.$nextTick(() => this.demoRemeasure()) },
  },
  computed: {
    /** 顶栏右侧演示按钮列表 (跟 Vue 3 demo src/App.vue:328-369 同款) */
    demoBarItems(): Array<{ id: string; kind: 'btn' | 'select' | 'color'; label: string; title?: string; onClick?: () => void; getColor?: () => string; onColor?: (e: Event) => void; model?: string; options?: { value: string; label: string }[]; onSelect?: (v: string) => void }> {
      if (!(this.src || this.jsonItems)) return []
      const arr: any[] = []
      if (this.editMode) {
        arr.push(
          { id: 'edit-targets', kind: 'btn', label: this.appliedCount ? `可编辑 (${this.appliedCount})` : '设置可编辑', title: '白名单模式: 点选要可编辑的格 / 行 / 列, 应用后只这些可编辑', onClick: this.openEditTargetsDialog },
          { id: 'highlight-readonly', kind: 'btn', label: this.highlightReadOnly ? '✓ 高亮只读' : '高亮只读', title: '把只读格套浅灰底', onClick: this.toggleHighlightReadOnly },
          { id: 'bold', kind: 'btn', label: 'B 加粗选区', title: '给选区加粗', onClick: this.boldSel },
          { id: 'merge', kind: 'btn', label: '合并', title: '合并选区', onClick: this.mergeSel },
          { id: 'unmerge', kind: 'btn', label: '拆分', title: '拆分选区', onClick: this.unmergeSel },
          { id: 'fill', kind: 'color', label: '背景', title: '背景填充色', getColor: this.getFill, onColor: this.setFill },
          { id: 'font', kind: 'color', label: '字体', title: '字体颜色', getColor: this.getFont, onColor: this.setFont },
          { id: 'clear-fill', kind: 'btn', label: '清除填充', onClick: this.clearFill },
          { id: 'embed-all', kind: 'btn', label: '整表嵌入', title: 'WPS 浮动→嵌入(DISPIMG)', onClick: this.embedAll },
          { id: 'cell-to-float', kind: 'btn', label: '格→图', title: '内嵌图→浮动图', onClick: this.cellToFloat },
          { id: 'ins-row', kind: 'btn', label: '＋行', title: '选区上方插入行', onClick: this.insRow },
          { id: 'del-row', kind: 'btn', label: '－行', title: '删除选区行', onClick: this.delRow },
          { id: 'paste-config', kind: 'btn', label: '⚙ 粘贴配置', title: '配置 Ctrl+V/右键粘贴行为(覆盖/合并/仅值、目标合并清除、列宽…)', onClick: () => this.v()?.openPasteConfigDialog() },
        )
      }
      arr.push(
        { id: 'pdf-watermark', kind: 'btn', label: 'PDF(页码+水印)', title: '演示 beforeRenderPage 钩子', onClick: this.exportPdfWithWatermark },
        { id: 'sheet-json', kind: 'btn', label: '数据→JSON', title: '演示数据读取 API getSheetJSON', onClick: this.showSheetJSON },
        { id: 'jump-last-row', kind: 'btn', label: '跳到末行', title: '演示 scrollToCell(row,col,{select:true}) 导航 API', onClick: this.jumpToLastRow },
        { id: 'fit', kind: 'select', label: '贴合', title: 'WPS 内嵌图贴合方式', model: this.cellImageFit, options: [
          { value: 'contain', label: 'contain 等比(同 WPS)' }, { value: 'fill', label: 'fill 铺满' }, { value: 'cover', label: 'cover 裁剪' },
        ], onSelect: (v: string) => { this.cellImageFit = v as any } },
        { id: 'dl-xlsx', kind: 'btn', label: '↓XLSX', title: '导出 .xlsx', onClick: this.downloadXlsx },
        { id: 'dl-csv', kind: 'btn', label: '↓CSV', title: '导出 .csv', onClick: this.downloadCsv },
        { id: 'dl-json', kind: 'btn', label: '↓JSON', title: '导出 .json', onClick: this.downloadJson },
      )
      return arr
    },
    demoVisibleCount(): number {
      const cw = this.demoBarContentW
      const w = this.demoItemWidths
      const items = this.demoBarItems
      const MORE_W = 64, GAP = 6
      if (!cw || w.length !== items.length) return items.length
      let sum = 0, fitsAll = true
      for (let i = 0; i < items.length; i++) {
        sum += w[i] + GAP
        if (sum > cw) { fitsAll = false; break }
      }
      if (fitsAll) return items.length
      let s = MORE_W, n = 0
      for (let i = 0; i < items.length; i++) {
        s += w[i] + GAP
        if (s > cw) break
        n++
      }
      return Math.max(0, n)
    },
    demoVisibleItems(): any[] { return this.demoBarItems.slice(0, this.demoVisibleCount) },
    demoOverflowItems(): any[] { return this.demoBarItems.slice(this.demoVisibleCount) },
    editableTargetsCount(): number {
      return Object.keys(this.editTargetsCells).length + Object.keys(this.editTargetsRows).length + Object.keys(this.editTargetsCols).length
    },
    appliedCount(): number {
      return (this.editableTargetsApplied as any[] | undefined)?.length ?? 0
    },
  },
  methods: {
    onFileInput(e: Event) {
      const f = (e.target as HTMLInputElement).files?.[0]
      if (!f) return
      this.fileName = f.name
      this.jsonItems = null
      ;(this as any).src = f
    },
    onDrop(e: DragEvent) {
      this.dragOver = false
      const f = e.dataTransfer?.files?.[0]
      if (!f) return
      this.fileName = f.name
      this.jsonItems = null
      ;(this as any).src = f
    },
    loadSample() {
      this.jsonItems = null
      ;(this as any).src = '/sample.xlsx'
      this.fileName = 'sample.xlsx'
    },
    loadJsonSample() {
      // 跟 Vue 3 demo 同款: 不传 src, 只给 :workbook=对象数组 → 默认渲染
      ;(this as any).src = undefined
      this.fileName = '订单数据'
      this.jsonItems = [
        { name: '笔记本电脑', price: 5999, qty: 1, amount: 5999, note: '商务款' },
        { name: '机械键盘', price: 399, qty: 2, amount: 798, note: '青轴' },
        { name: '显示器', price: 1299, qty: 2, amount: 2598, note: '27寸 2K' },
        { name: '鼠标', price: 89, qty: 5, amount: 445, note: '无线' },
        { name: '耳机', price: 599, qty: 3, amount: 1797, note: '降噪' },
      ]
    },
    async exportPdfWithWatermark() {
      const viewer = (this.$refs.viewer as any)
      if (!viewer) return
      try {
        await viewer.downloadPdf({
          target: 'all',
          beforeRenderPage: (ctx: any) => {
            const { doc, pageIndex, pageCount, pageWidth, pageHeight, margin, sheetName } = ctx
            doc.setFontSize(9); doc.setTextColor(120)
            doc.text(`${sheetName}`, margin.left, pageHeight - 5)
            doc.text(`第 ${pageIndex + 1} / ${pageCount} 页`, pageWidth - margin.right, pageHeight - 5, { align: 'right' })
            doc.setFontSize(56); doc.setTextColor(230)
            doc.text('PREVIEW', pageWidth / 2, pageHeight / 2, { align: 'center', angle: 30 })
          },
        })
        this.lastEvent = '已导出 PDF (全部表 + 页码 + 水印)'
      } catch (e) {
        this.lastEvent = '导出失败: ' + (e as Error).message
      }
    },
    showSheetJSON() {
      const viewer = (this.$refs.viewer as any)
      if (!viewer) return
      const json = viewer.getSheetJSON({ headerRow: 1 })
      navigator.clipboard?.writeText(JSON.stringify(json, null, 2)).catch(() => {})
      this.lastEvent = `[数据] ${json.length} 行已复制为 JSON · 首行: ${JSON.stringify(json[0] ?? {})}`.slice(0, 140)
    },
    jumpToLastRow() {
      const viewer = (this.$refs.viewer as any)
      const wb = viewer?.getWorkbook?.()
      if (!viewer || !wb) return
      const sheet = wb.sheets[viewer.getActiveSheet()]
      const row = Math.max(0, sheet.dimension.rows - 1)
      viewer.scrollToCell(row, 0, { select: true })
      this.lastEvent = `已跳到末行 A${row + 1}`
    },
    downloadXlsx() { (this.$refs.viewer as any)?.downloadXlsx() },
    downloadCsv() { (this.$refs.viewer as any)?.downloadCsv() },
    downloadJson() { (this.$refs.viewer as any)?.downloadJson() },
    // ---- 编辑模式演示按钮 (跟 Vue 3 demo src/App.vue:347-367 对齐) ----
    v() { return this.$refs.viewer as any },
    boldSel() { const s = this.v()?.getSelection(); if (s) this.v()?.setStyle(s, { font: { bold: true } }) },
    mergeSel() { const s = this.v()?.getSelection(); if (s) this.v()?.mergeCells(s) },
    unmergeSel() { const s = this.v()?.getSelection(); if (s) this.v()?.unmergeCells(s) },
    getFill() { return this.v()?.getActiveFillColor() ?? '#FFFFFF' },
    getFont() { return this.v()?.getActiveFontColor() ?? '#000000' },
    setFill(e: Event) { this.v()?.setSelectionFill((e.target as HTMLInputElement).value) },
    setFont(e: Event) { this.v()?.setSelectionFontColor((e.target as HTMLInputElement).value) },
    clearFill() { this.v()?.setSelectionFill(null) },
    embedAll() { const n = this.v()?.convertAllImagesToCells() ?? 0; if (!n) this.lastEvent = '没有可嵌入的浮动图' },
    cellToFloat() { const s = this.v()?.getSelection(); if (s) this.v()?.convertCellImageToFloat(s.top, s.left) },
    insRow() { const s = this.v()?.getSelection(); if (s) this.v()?.insertRows(s.top, 1) },
    delRow() { const s = this.v()?.getSelection(); if (s) this.v()?.deleteRows(s.top, s.bottom - s.top + 1) },
    toggleHighlightReadOnly() { this.highlightReadOnly = !this.highlightReadOnly },
    // ---- 设置可编辑 dialog (白名单 API) ----
    openEditTargetsDialog() {
      this.editTargetsCells = {}
      this.editTargetsRows = {}
      this.editTargetsCols = {}
      const applied = this.editableTargetsApplied as any[] | undefined
      for (const t of applied ?? []) {
        if ('top' in t) continue
        if ('row' in t && 'col' in t && typeof t.col === 'number') this.$set(this.editTargetsCells, `${t.row}:${t.col}`, true)
        else if ('row' in t && typeof t.row === 'number') this.$set(this.editTargetsRows, t.row, true)
        else if ('col' in t && typeof t.col === 'number') this.$set(this.editTargetsCols, t.col, true)
      }
      this.editTargetsDialogOpen = true
    },
    toggleEditTargetCell(r: number, c: number) {
      const k = `${r}:${c}`
      if (this.editTargetsCells[k]) this.$delete(this.editTargetsCells, k)
      else this.$set(this.editTargetsCells, k, true)
    },
    toggleEditTargetRow(r: number) {
      if (this.editTargetsRows[r]) this.$delete(this.editTargetsRows, r)
      else this.$set(this.editTargetsRows, r, true)
    },
    toggleEditTargetCol(c: number) {
      if (this.editTargetsCols[c]) this.$delete(this.editTargetsCols, c)
      else this.$set(this.editTargetsCols, c, true)
    },
    isCellInDraft(r: number, c: number): boolean {
      return !!this.editTargetsCells[`${r}:${c}`] || !!this.editTargetsRows[r] || !!this.editTargetsCols[c]
    },
    previewCellText(r: number, c: number): string {
      const v = (this.v() as any)?.getCellText(r, c) ?? ''
      return v.length > 6 ? v.slice(0, 6) + '…' : v
    },
    colLetter(c: number): string {
      let s = ''; let n = c
      while (true) { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; if (n < 0) break }
      return s
    },
    applyEditTargets() {
      const arr: any[] = []
      for (const r of Object.keys(this.editTargetsRows)) arr.push({ row: Number(r) })
      for (const c of Object.keys(this.editTargetsCols)) arr.push({ col: Number(c) })
      for (const k of Object.keys(this.editTargetsCells)) {
        const [r, c] = k.split(':').map(Number)
        if (this.editTargetsRows[r] || this.editTargetsCols[c]) continue
        arr.push({ row: r, col: c })
      }
      this.editableTargetsApplied = arr
      this.editTargetsDialogOpen = false
      this.lastEvent = `[白名单] ${arr.length} 项 target 已应用; 其它格只读`
    },
    clearEditTargets() {
      this.editableTargetsApplied = undefined
      this.editTargetsDialogOpen = false
      this.lastEvent = '[白名单] 已关闭, 恢复默认 (全可编辑)'
    },
    demoRemeasure() {
      const bar = this.$refs.demoBarEl as HTMLElement | undefined
      const measure = this.$refs.demoMeasureEl as HTMLElement | undefined
      if (!measure || !bar) return
      this.demoItemWidths = Array.from(measure.children).map((c) => (c as HTMLElement).offsetWidth)
      const fixed = Array.from(bar.children).find((c) => (c as HTMLElement).classList.contains('app-bar-fixed')) as HTMLElement | undefined
      const fixedW = fixed ? fixed.getBoundingClientRect().width : 0
      this.demoBarContentW = Math.max(0, bar.clientWidth - fixedW - 24)
    },
    onRendered() { this.lastEvent = '✓ 渲染完成' },
    onError(msg: string) { this.lastEvent = '⚠ 错误: ' + msg },
    onCellClick(p: { row: number; col: number; text: string }) { this.lastEvent = `点击 R${p.row + 1}C${p.col + 1}: ${p.text}` },
    onSelectionChange(p: { range: { top: number; left: number; bottom: number; right: number } }) {
      const r = p.range
      this.lastEvent = `选区 ${r.top + 1},${r.left + 1} → ${r.bottom + 1},${r.right + 1}`
    },
    onCellChange(p: { before: { text: string }; after: { text: string }; source: string }) {
      this.lastEvent = `[${p.source}] "${p.before.text}" → "${p.after.text}"`
    },
    onSheetChange(p: { index: number; name: string }) { this.lastEvent = `切到 sheet[${p.index}] ${p.name}` },
    onPermissionDenied(p: { reason: string; cells: { row: number; col: number }[] }) {
      this.lastEvent = `🚫 [${p.reason}] 拒绝 ${p.cells.length} 个格`
    },
  },
  template: `
    <div style="display:flex;flex-direction:column;height:100vh" :class="{ dragging: dragOver }"
         @dragover.prevent="dragOver = true" @dragleave.prevent="dragOver = false" @drop.prevent="onDrop">
      <header ref="demoBarEl" class="app-bar">
        <div class="app-bar-fixed">
          <strong>OOXML Excel 预览器</strong>
          <span class="sub">Vue 2 · Canvas 高保真</span>
          <label class="file-btn">
            选择 .xlsx
            <input type="file" accept=".xlsx,.xlsm" @change="onFileInput" hidden />
          </label>
          <button class="sample-btn" @click="loadSample">加载示例</button>
          <button class="sample-btn" @click="loadJsonSample" title="加载一个 JSON 数据源演示;然后用工具栏「模板」导入 .xlsx 看模板效果">JSON 示例</button>
          <label v-if="src || jsonItems" class="edit-toggle" title="开启编辑模式">
            <input type="checkbox" v-model="editMode" /> 编辑模式
          </label>
        </div>
        <div class="grow"></div>
        <!-- 隐藏测量行: 量每项实际宽度 -->
        <div ref="demoMeasureEl" class="app-bar-measure" aria-hidden="true">
          <template v-for="it in demoBarItems">
            <button v-if="it.kind === 'btn'" :key="'m' + it.id" class="sample-btn">{{ it.label }}</button>
            <label v-else-if="it.kind === 'color'" :key="'m' + it.id" class="sample-label">{{ it.label }}<input type="color" /></label>
            <label v-else :key="'m' + it.id" class="sample-label">{{ it.label }}<select><option v-for="o in it.options" :key="o.value" :value="o.value">{{ o.label }}</option></select></label>
          </template>
        </div>
        <!-- 可见演示按钮 -->
        <template v-for="it in demoVisibleItems">
          <button v-if="it.kind === 'btn'" :key="it.id" class="sample-btn" :title="it.title" @click="it.onClick && it.onClick()">{{ it.label }}</button>
          <label v-else-if="it.kind === 'color'" :key="it.id" class="sample-label" :title="it.title">{{ it.label }}<input type="color" :value="it.getColor && it.getColor()" @input="it.onColor && it.onColor($event)" /></label>
          <label v-else :key="it.id" class="sample-label" :title="it.title">{{ it.label }}<select :value="it.model" @change="it.onSelect && it.onSelect($event.target.value)"><option v-for="o in it.options" :key="o.value" :value="o.value">{{ o.label }}</option></select></label>
        </template>
        <!-- 更多溢出 popover -->
        <div v-if="demoOverflowItems.length" class="more-wrap">
          <button class="sample-btn more-btn" :class="{ open: demoMoreOpen }" title="更多" @click="demoMoreOpen = !demoMoreOpen">⋯ 更多</button>
          <div v-if="demoMoreOpen" class="more-pop">
            <template v-for="it in demoOverflowItems">
              <button v-if="it.kind === 'btn'" :key="'o' + it.id" class="more-row" :title="it.title" @click="(it.onClick && it.onClick()); demoMoreOpen = false">{{ it.label }}</button>
              <label v-else-if="it.kind === 'color'" :key="'o' + it.id" class="more-row" :title="it.title">{{ it.label }}<input type="color" :value="it.getColor && it.getColor()" @input="it.onColor && it.onColor($event)" /></label>
              <label v-else :key="'o' + it.id" class="more-row" :title="it.title">{{ it.label }}<select :value="it.model" @change="(it.onSelect && it.onSelect($event.target.value)); demoMoreOpen = false"><option v-for="o in it.options" :key="o.value" :value="o.value">{{ o.label }}</option></select></label>
            </template>
          </div>
        </div>
      </header>
      <main style="position:relative;flex:1 1 auto;min-height:0">
        <ExcelViewer
          ref="viewer"
          :src="src"
          :workbook="jsonItems"
          :file-name="fileName"
          :cell-image-fit="cellImageFit"
          :editable="editMode"
          :pivot-table="true"
          :conditional-format="true"
          :recalc="editMode"
          :editable-targets="editableTargetsApplied"
          :read-only-cell-style="highlightReadOnly"
          :toolbar="toolbarItems"
          @rendered="onRendered"
          @error="onError"
          @cell-click="onCellClick"
          @selection-change="onSelectionChange"
          @cell-change="onCellChange"
          @sheet-change="onSheetChange"
          @permission-denied="onPermissionDenied"
        />
        <div v-if="lastEvent" style="position:absolute;left:12px;bottom:12px;background:rgba(0,0,0,.78);color:#fff;font-size:12px;padding:5px 10px;border-radius:5px;pointer-events:none;z-index:9">{{ lastEvent }}</div>
      </main>
      <div v-if="editTargetsDialogOpen" class="edit-targets-overlay" @click.self="editTargetsDialogOpen = false">
        <div class="edit-targets-dialog">
          <header>
            <h3>设置可编辑单元格 (白名单)</h3>
            <p class="hint">
              点击单元格 = 该格可编辑;点击列标题 (A/B/C…) = 整列可编辑;点击行号 = 整行可编辑.
              应用后,只有勾选的位置可编辑,其它全部只读. 关闭白名单 = 恢复默认 (整表可编辑).
            </p>
          </header>
          <div class="edit-targets-grid">
            <table>
              <thead>
                <tr>
                  <th class="corner">#</th>
                  <th v-for="c in 8" :key="'h' + c"
                      :class="{ picked: !!editTargetsCols[c - 1] }"
                      @click="toggleEditTargetCol(c - 1)"
                      :title="'整列 ' + colLetter(c - 1) + ' 可编辑'">{{ colLetter(c - 1) }}</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="r in 12" :key="'r' + r">
                  <th :class="{ picked: !!editTargetsRows[r - 1] }"
                      @click="toggleEditTargetRow(r - 1)"
                      :title="'整行 ' + r + ' 可编辑'">{{ r }}</th>
                  <td v-for="c in 8" :key="'c' + r + ',' + c"
                      :class="{ picked: isCellInDraft(r - 1, c - 1), 'row-col-hit': !!editTargetsRows[r - 1] || !!editTargetsCols[c - 1] }"
                      @click="toggleEditTargetCell(r - 1, c - 1)"
                      :title="'R' + r + 'C' + c">{{ previewCellText(r - 1, c - 1) || '·' }}</td>
                </tr>
              </tbody>
            </table>
          </div>
          <footer>
            <span class="count-hint">
              已选: {{ Object.keys(editTargetsCells).length }} 单格 / {{ Object.keys(editTargetsRows).length }} 整行 / {{ Object.keys(editTargetsCols).length }} 整列
            </span>
            <button class="dlg-btn ghost" @click="editTargetsDialogOpen = false">取消</button>
            <button class="dlg-btn ghost" @click="clearEditTargets" title="移除白名单, 恢复默认 (全可编辑)">关闭白名单</button>
            <button class="dlg-btn primary" @click="applyEditTargets">应用</button>
          </footer>
        </div>
      </div>
    </div>
  `,
})
