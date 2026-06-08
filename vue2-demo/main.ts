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
      editMode: false,
      highlightReadOnly: false,
      cellImageFit: 'contain' as 'contain' | 'fill' | 'cover',
      lastEvent: '',
      // 跟 Vue 3 demo (src/App.vue:513) 同款完整工具栏配置
      toolbarItems: ['find', 'filter', 'clear-filter', 'separator', 'copy', 'wrap-text', 'image-tools', 'freeze', 'separator', 'template', 'separator', 'zoom', 'export'],
    }
  },
  methods: {
    onFileInput(e: Event) {
      const f = (e.target as HTMLInputElement).files?.[0]
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
    <div style="display:flex;flex-direction:column;height:100vh">
      <header class="app-bar">
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
        <template v-if="(src || jsonItems) && editMode">
          <button class="sample-btn" @click="toggleHighlightReadOnly" :title="highlightReadOnly ? '关闭只读高亮' : '把只读格套浅灰底'">{{ highlightReadOnly ? '✓ 高亮只读' : '高亮只读' }}</button>
          <button class="sample-btn" @click="boldSel" title="给选区加粗">B 加粗选区</button>
          <button class="sample-btn" @click="mergeSel" title="合并选区">合并</button>
          <button class="sample-btn" @click="unmergeSel" title="拆分选区">拆分</button>
          <label class="sample-label" title="背景填充色(回显 + 改选区)">背景<input type="color" :value="getFill()" @input="setFill" /></label>
          <label class="sample-label" title="字体颜色(回显 + 改选区)">字体<input type="color" :value="getFont()" @input="setFont" /></label>
          <button class="sample-btn" @click="clearFill" title="清除背景填充">清除填充</button>
          <button class="sample-btn" @click="embedAll" title="整表浮动图就近嵌入(WPS 浮动→嵌入/DISPIMG)">整表嵌入</button>
          <button class="sample-btn" @click="cellToFloat" title="选中格的内嵌图拎成浮动图">格→图</button>
          <button class="sample-btn" @click="insRow" title="选区上方插入行">＋行</button>
          <button class="sample-btn" @click="delRow" title="删除选区行">－行</button>
        </template>
        <template v-if="src || jsonItems">
          <button class="sample-btn" @click="exportPdfWithWatermark" title="演示 beforeRenderPage 钩子">PDF(页码+水印)</button>
          <button class="sample-btn" @click="showSheetJSON" title="演示数据读取 API getSheetJSON">数据→JSON</button>
          <label class="sample-label" title="WPS 内嵌图贴合方式">
            贴合
            <select v-model="cellImageFit">
              <option value="contain">contain 等比(同 WPS)</option>
              <option value="fill">fill 铺满</option>
              <option value="cover">cover 裁剪</option>
            </select>
          </label>
          <button class="sample-btn" @click="downloadXlsx" title="导出 .xlsx">↓XLSX</button>
          <button class="sample-btn" @click="downloadCsv" title="导出 .csv">↓CSV</button>
          <button class="sample-btn" @click="downloadJson" title="导出 .json">↓JSON</button>
        </template>
      </header>
      <main style="position:relative;flex:1 1 auto;min-height:0">
        <ExcelViewer
          ref="viewer"
          :src="src"
          :workbook="jsonItems"
          :file-name="fileName"
          :cell-image-fit="cellImageFit"
          :editable="editMode"
          :recalc="editMode"
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
    </div>
  `,
})
