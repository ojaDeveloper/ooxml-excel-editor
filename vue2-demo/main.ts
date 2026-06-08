// Vue 2 demo 入口. 显式从 'vue2' alias (vue@2.7) 导入, 避免跟 Vue 3 root 冲突.
import Vue from 'vue2'
import ExcelViewer from '../src/vue2/ExcelViewer'

new Vue({
  el: '#app',
  components: { ExcelViewer },
  data() {
    return {
      src: undefined as string | undefined,
      editMode: false,
      lastEvent: '',
      fileName: '',
    }
  },
  methods: {
    onFileInput(e: Event) {
      const f = (e.target as HTMLInputElement).files?.[0]
      if (!f) return
      this.fileName = f.name
      // Vue 2 不能直接 set 这种类型, 用 $set 或 cast
      ;(this as any).src = f
    },
    loadSample() {
      ;(this as any).src = '/sample.xlsx'
      this.fileName = 'sample.xlsx'
    },
    onCellClick(p: { row: number; col: number; text: string }) {
      this.lastEvent = `点击 R${p.row + 1}C${p.col + 1}: ${p.text}`
    },
    onSelectionChange(p: { range: { top: number; left: number; bottom: number; right: number } }) {
      const r = p.range
      this.lastEvent = `选区 ${r.top + 1},${r.left + 1} → ${r.bottom + 1},${r.right + 1}`
    },
    onRendered() {
      this.lastEvent = '渲染完成'
    },
    onError(msg: string) {
      this.lastEvent = '错误: ' + msg
    },
    downloadXlsx() {
      ;(this.$refs.viewer as any)?.downloadXlsx?.()
    },
    downloadPdf() {
      ;(this.$refs.viewer as any)?.downloadPdf?.()
    },
  },
  template: `
    <div style="display:flex;flex-direction:column;height:100vh">
      <header style="display:flex;align-items:center;gap:8px;height:48px;padding:0 16px;background:#21a366;color:#fff">
        <strong>OOXML Excel — Vue 2 Demo</strong>
        <span style="font-size:12px;opacity:.85">render-function 实现</span>
        <label style="background:rgba(255,255,255,.18);color:#fff;border:1px solid rgba(255,255,255,.4);padding:6px 14px;border-radius:5px;cursor:pointer;font-size:13px">
          选择 .xlsx
          <input type="file" accept=".xlsx,.xlsm" @change="onFileInput" hidden />
        </label>
        <button @click="loadSample" style="background:rgba(255,255,255,.18);color:#fff;border:1px solid rgba(255,255,255,.4);padding:6px 14px;border-radius:5px;cursor:pointer;font-size:13px">加载示例</button>
        <label style="font-size:13px"><input type="checkbox" v-model="editMode" /> 编辑模式</label>
        <button @click="downloadXlsx" style="background:rgba(255,255,255,.18);color:#fff;border:1px solid rgba(255,255,255,.4);padding:6px 14px;border-radius:5px;cursor:pointer;font-size:13px;margin-left:auto">↓XLSX</button>
        <button @click="downloadPdf" style="background:rgba(255,255,255,.18);color:#fff;border:1px solid rgba(255,255,255,.4);padding:6px 14px;border-radius:5px;cursor:pointer;font-size:13px">↓PDF</button>
      </header>
      <main style="position:relative;flex:1 1 auto;min-height:0">
        <ExcelViewer
          ref="viewer"
          :src="src"
          :file-name="fileName"
          :editable="editMode"
          @rendered="onRendered"
          @error="onError"
          @cell-click="onCellClick"
          @selection-change="onSelectionChange"
        />
        <div v-if="lastEvent" style="position:absolute;left:12px;bottom:12px;background:rgba(0,0,0,.78);color:#fff;font-size:12px;padding:5px 10px;border-radius:5px;pointer-events:none">{{ lastEvent }}</div>
      </main>
    </div>
  `,
})
