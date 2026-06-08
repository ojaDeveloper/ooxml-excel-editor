/**
 * Vue 2 demo. 显式从 'vue2' alias (vue@2.7) 导入, 避免跟 Vue 3 root 冲突.
 * Vue 2 壳内置了 header / action toolbar / 公式栏 / sheet 标签 / 查找 — 跟 Vue 3 壳同构,
 * demo 只用一个最薄的外壳挂选文件按钮 + 加载示例.
 */
import Vue from 'vue2'
import ExcelViewer from '../src/vue2/ExcelViewer'
import '../src/vue2/excel-viewer.css'

new Vue({
  el: '#app',
  components: { ExcelViewer },
  data() {
    return {
      src: undefined as string | File | undefined,
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
      ;(this as any).src = f
    },
    loadSample() {
      ;(this as any).src = '/sample.xlsx'
      this.fileName = 'sample.xlsx'
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
    <div style="display:flex;flex-direction:column;height:100vh">
      <header style="display:flex;align-items:center;gap:10px;height:48px;padding:0 16px;background:#21a366;color:#fff">
        <strong>OOXML Excel</strong>
        <span style="font-size:12px;opacity:.85">Vue 2 壳 (Composition API · render-function · 共享 core)</span>
        <label style="background:rgba(255,255,255,.18);color:#fff;border:1px solid rgba(255,255,255,.4);padding:6px 14px;border-radius:5px;cursor:pointer;font-size:13px">
          选择 .xlsx
          <input type="file" accept=".xlsx,.xlsm" @change="onFileInput" hidden />
        </label>
        <button @click="loadSample" style="background:rgba(255,255,255,.18);color:#fff;border:1px solid rgba(255,255,255,.4);padding:6px 14px;border-radius:5px;cursor:pointer;font-size:13px">加载示例</button>
        <label style="font-size:13px"><input type="checkbox" v-model="editMode" /> 编辑模式</label>
      </header>
      <main style="position:relative;flex:1 1 auto;min-height:0">
        <ExcelViewer
          ref="viewer"
          :src="src"
          :file-name="fileName"
          :editable="editMode"
          :recalc="editMode"
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
