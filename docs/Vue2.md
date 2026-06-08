# Vue 2 兼容子入口(1.3.0)

让 **Vue 2.6.x / 2.7+** 项目用本库,跟 Vue 3 / React 壳共享同一份 framework-agnostic core,**~100% 功能对齐**。

---

## 1. 安装 + 引入

### Vue 2.7+ (推荐)

```bash
npm i ooxml-excel-editor vue@2.7 exceljs
npm i @vue/composition-api    # 必装 (即使 Vue 2.7 内置 Composition API)
# 可选 peer (按需): echarts, hyperformula, jspdf
```

Vue 2.7+ 内置 Composition API, `@vue/composition-api` 检测到后**自动 re-export 内置 API** (plugin install 为 noop), **无需 `Vue.use()` 注册**.

### Vue 2.6.x

```bash
npm i ooxml-excel-editor vue@2.6 @vue/composition-api exceljs
```

```js
// main.js — 必须显式注册 @vue/composition-api plugin
import Vue from 'vue'
import VueCompositionAPI from '@vue/composition-api'
Vue.use(VueCompositionAPI)
```

---

```html
<!-- 你的项目 (Vue 2.6.x 或 2.7+) -->
<template>
  <ExcelViewer
    ref="viewer"
    :src="src"
    :editable="editing"
    :file-name="fileName"
    @rendered="onRendered"
    @cell-change="onCellChange"
    @selection-change="onSelection"
  />
</template>

<script>
import ExcelViewer from 'ooxml-excel-editor/vue2'
import 'ooxml-excel-editor/style.css'
import 'ooxml-excel-editor/vue2.css'

export default {
  components: { ExcelViewer },
  data: () => ({ src: null, editing: false, fileName: '' }),
  methods: {
    download() { this.$refs.viewer.downloadXlsx() },
    onRendered(wb) { console.log('sheets:', wb.sheets.length) },
    onCellChange(p) { console.log(p.before.text, '→', p.after.text) },
    onSelection(p) { console.log(p.range, p.active) },
  },
}
</script>
```

### 为什么要装 `@vue/composition-api`?

Vue 2 壳代码统一从 `@vue/composition-api` import Composition API (`ref` / `computed` / `watch` / `onMounted` / `defineComponent` / ...). 这样:
- **Vue 2.6.x**: 通过 plugin 注入 Composition API (它本来没有)
- **Vue 2.7+**: plugin 检测到内置 API → re-export 它, plugin install 是 noop

一份代码同时支持 2.6 + 2.7,你只需正确装 + 2.6 用户加 `Vue.use(VueCompositionAPI)` 一行.

---

## 2. 架构说明

### 跟 Vue 3 / React 壳的关系

| 入口 | 文件 | 体积 | 实现 |
|---|---|---|---|
| `ooxml-excel-editor` | `dist/index.js` | ~71 KB | Vue 3 (SFC) |
| `ooxml-excel-editor/react` | `dist/react.js` | ~50 KB | React (.tsx + hook) |
| `ooxml-excel-editor/vue2` | `dist/vue2.js` | ~396 KB | Vue 2.7 (.ts + Composition API + render function) |
| `ooxml-excel-editor/core` | `dist/core.js` | ~2 KB | 框架无关 (动态加载 chunks/) |

三种壳共享同一份 `chunks/plugin-overlay-*.js` (含核心引擎 + ViewerController + 渲染 + 编辑 + 导出 + 公式重算)。

**Vue 2 入口体积较大的原因**: 跟 Vue 3 / React 入口共享 chunk 需要在同一次 lib build 内,但 Vue 2 跟 Vue 3 同时存在编译冲突 (`@vue/compiler-sfc` 解析路径), 所以 Vue 2 用 **独立 build pass** (`vite build --mode lib-vue2`), 不共享 chunks. 后续可能通过自定义 rollup 插件优化,暂时接受 ~290KB 的额外体积。

### 为什么是 `.ts + render function` 而非 `.vue` SFC

项目同时存在 `vue@3.5` (Vue 3 壳的 peer) 和 `vue@2.7` (alias `vue2`), 导致:

```
@vitejs/plugin-vue2 加载 @vue/compiler-sfc → npm 解析到 Vue 3 的版本
→ 编译 Vue 2 SFC 用 Vue 3 编译器 → "currentInput.slice is not a function"
```

绕开方案: Vue 2 壳完全用 **render function** 实现, 不走 SFC 编译路径。Vue 2.7 内置 Composition API + `h()`, 跟 React 壳的 .tsx + hook 工程上几乎同构。

---

## 3. API 对齐

### Props (跟 Vue 3 壳完全对齐, 28 项)

| 类别 | Props |
|---|---|
| **数据源** | `src` / `workbook` / `jsonOptions` / `templateFile` / `templateName` / `fileName` |
| **主题** | `theme` / `cellStyle` / `cellImageFit` / `readOnlyCellStyle` |
| **交互** | `imageLightbox` / `openLinks` / `contextMenu` / `toolbar` |
| **编辑** | `editable` / `cellReadOnly` / `readOnlyRanges` / `editableTargets` / `strictDimensions` / `editor` / `recalc` / `formulaEngine` |
| **数据钩子** | `transformModel` / `plugins` |

### Events (15 项)

```
rendered / error / progress / cell-click / cell-dblclick / selection-change /
sheet-change / hyperlink-click / cell-change / edit-start / edit-commit /
dim-change / dirty-change / image-change / struct-change / permission-denied /
before-context-menu / context-menu
```

### 命令式 API (`this.$refs.viewer`, 全套 80+ 方法)

跟 Vue 3 `ViewerApi` / React `ExcelViewerHandle` 完全对齐:

```ts
// 数据
viewer.load(src) / getWorkbook() / getActiveSheet() / setActiveSheet(i) /
       getSelection() / setSelection(range) / rectOf(r,c) / rectOfRange(range)

// 数据读取 (跟 Vue 3 同形)
viewer.getCellValue(r,c) / getCellText(r,c) / getSheetData() /
       getSheetJSON() / getRangeData(range)

// 编辑 + 命令栈
viewer.editCell(r,c,v) / editRange(range, values) / clearRange(range) /
       setStyle(range, patch) / mergeCells / unmergeCells / pasteText /
       insertRows / deleteRows / insertCols / deleteCols / undo / redo

// 编辑权限 (1.2.0 起)
viewer.isCellEditable(r,c) / setEditableTargets(targets) / getEditableTargets() /
       isDirty() / resetToOriginal()

// 尺寸 (1.2.0 多形态)
viewer.setColumnWidth(target, w) / setRowHeight(target, h) /
       autoFitColumns(target?) / autoFitRows(target?) /
       resetColumnWidth(target) / resetRowHeight(target)

// 编辑状态机
viewer.beginEdit(r,c) / cancelEdit() / isEditing() / getEditingCell() /
       getCellSnapshot(r,c) / inspectCell(r,c)

// 公式栏
viewer.getCellEditString() / canEditActiveCell() /
       commitActiveCellValue(value, move)

// 图片
viewer.getImages() / addImage / removeImage / moveImage / resizeImage /
       getCellImages() / getCellImageAt(r,c) / openImageLightbox /
       convertImageToCell / convertAllImagesToCells / convertCellImageToFloat

// 样式 + 格式化
viewer.getActiveFillColor() / getActiveFontColor() /
       setSelectionFill / setSelectionFontColor /
       getSelectionWrapState() / toggleWrapTextOnSelection()

// 导出 + 打印
viewer.exportImage(opts?) / downloadImage / exportPdf / downloadPdf / print /
       exportXlsx / downloadXlsx / exportJson / downloadJson / exportCsv / downloadCsv

// 右键菜单 (Plan C)
viewer.openContextMenu(x, y, items?) / closeContextMenu()
```

### Vue 2 特定

- **`v-model` 单值** (跟 Vue 3 多 v-model 区别): Vue 2 自定义组件 v-model 只支持单 prop, 本组件没有 v-model, 不影响
- **`$refs.viewer` 拿命令式 API**: 跟 Vue 3 `ref` 一致, 全套方法都挂在 `$refs.viewer` 上
- **作用域插槽语法**: Vue 2 `slot-scope` 跟 Vue 3 `v-slot` 写法不同; 但本 MVP 还没暴露插槽 (后续补)

---

## 4. 内置 UI 元素 (1.3.0)

Vue 2 壳自带跟 Vue 3 同款的 UI:

- **顶部工具栏**: 文件名 + 模板名 + 表数 + 缩放下拉 + 导出 PNG/PDF/XLSX 按钮
- **Action 工具栏**: 查找 / 筛选 / 清除筛选 / 复制 / 冻结 + 插件 toolbar items
- **公式栏**: textarea + auto-resize (跟 1.2.1 WPS 风格撑高一致)
- **Sheet 标签**: 多表切换 (1 表时隐藏)
- **查找条**: Ctrl+F 唤起,跟 Vue 3 同 keyboard binding
- **单元格编辑器**: WPS 长文本撑高 (跟 1.2.1 一致)
- **右键菜单**: Plan C 全套 (insert/delete/clear/merge 等)
- **图片放大灯箱**: 单击图 / 右键查看大图 / 下载原图

---

## 5. 已知限制 + 后续 (1.3.x 路线)

| 限制 | 现状 | 路线 |
|---|---|---|
| 包体积 | 396 KB (含内嵌 core) | 后续探索 rollup 多入口共享 chunk |
| 暂未暴露的 slot | header / toolbar / statusbar / overlay / export-progress | 1.3.x 补 (Vue 2 作用域插槽语法略不同) |
| 导出对话框 | 没内置 (Vue 3 壳有 `ExportDialog.vue`) | 用方调命令式 `downloadXlsx()` 等即可, 或自行实现 UI |
| `.d.ts` 类型声明 | `dist/vue2.d.ts` 未生成 (vue-tsc 不认 Vue 2 SFC; 本入口 .ts 后续可加) | 1.3.0 正式版补 |
| e2e 覆盖 | Playwright 多入口 vue2-demo 待配 | 1.3.0 正式版补 |

---

## 6. Vue 2 vs Vue 3 用方差异速查

| 维度 | Vue 3 | Vue 2 |
|---|---|---|
| 引入 | `import ExcelViewer from 'ooxml-excel-editor'` | `import ExcelViewer from 'ooxml-excel-editor/vue2'` |
| 全局 CSS | `import 'ooxml-excel-editor/style.css'` | 加 `+ 'ooxml-excel-editor/vue2.css'` |
| 注册 | 全局或本地 | `components: { ExcelViewer }` |
| 命令式访问 | `<ExcelViewer ref="v" />` + `v.value.foo()` | `<ExcelViewer ref="viewer" />` + `this.$refs.viewer.foo()` |
| 事件 | `@cell-change="..."` (同 Vue 2) | `@cell-change="..."` (同 Vue 3) |
| 插槽 | `<template #header="{ workbook }">` | (本 MVP 暂无, 后续补) |
| 插件 | `:plugins="[myPlugin]"` | 同 |

---

## ★ webpack 4 / vue-cli 4 兼容 (1.3.2+)

老打包器(webpack 4 / vue-cli 4)有几个独特限制, 1.3.2 已经处理:

### 1. 包根 stub (解决 webpack 4 不读 `package.json#exports`)

webpack 5+ / Vite / Rollup 都支持 `package.json#exports` 字段映射子路径. webpack 4 **不支持**, 解析 `'ooxml-excel-editor/vue2'` 会失败. 我们在包根放了 `vue2.js` / `react.js` / `core.js` stub re-export 到 dist, webpack 4 自动找得到:

```js
// webpack 4 (Vue CLI 4) — 推荐写法, 通用
import ExcelViewer from 'ooxml-excel-editor/vue2'
import 'ooxml-excel-editor/dist/style.css'  // CSS 走真实路径
import 'ooxml-excel-editor/dist/vue2.css'   // CSS 走真实路径

// 或者直接走真实路径, 任何打包器都识别
import ExcelViewer from 'ooxml-excel-editor/dist/vue2.js'
```

### 2. 语法降级 (ES2018, 解决 `??` / `?.` 老 webpack 4 SyntaxError)

dist 已 `target: 'es2018'`. 没有 `??` (空值合并 ES2020) / `?.` (可选链 ES2020). 保留 async/await + spread/rest (ES2017, webpack 4 OK). **无需 `transpileDependencies` 转译本包**.

### 3. 无 module worker / `import.meta.url` (库构建走 stub 主线程解析)

1.3.2 起 dist 全部走 `worker-client.stub.ts` (主线程跑解析), **不**用 `new Worker(new URL(..., import.meta.url))`. webpack 4 解析 dist 不会 SyntaxError, 运行也不缺 worker chunk.

### 4. 可选能力不放 peerDeps (避免 npm 7+ ERESOLVE)

`echarts` / `jspdf` / `hyperformula` **不在 peerDependencies**. 但 **webpack 4** 静态扫描 `await import('xxx')` 时如果 lib 没装会发 **warning** (不是 error, build 仍通过, 不调相应功能时也不会 runtime 报错). warning 格式:

```
* hyperformula in ./node_modules/ooxml-excel-editor/dist/vue2.js
* jspdf in ./node_modules/ooxml-excel-editor/dist/vue2.js
To install them, you can run: npm install --save hyperformula jspdf
```

#### 消除 warning 方案 A (推荐): webpack.config IgnorePlugin

`vue.config.js` 加 IgnorePlugin 让 webpack 跳过这 3 个 lib 的解析:

```js
// vue.config.js
const { IgnorePlugin } = require('webpack')

module.exports = {
  configureWebpack: {
    plugins: [
      new IgnorePlugin({ resourceRegExp: /^(echarts|hyperformula|jspdf)$/ }),
    ],
  },
}
```

用对应功能时再注释掉 + 装:
- 用图表: 注释掉 `echarts` + `npm i echarts`
- 用 PDF: 注释掉 `jspdf` + `npm i jspdf`
- 用编辑 + 公式重算: 注释掉 `hyperformula` + `npm i hyperformula`

#### 方案 B: 直接装 3 个 lib (一劳永逸)

```bash
npm i echarts jspdf hyperformula
```

webpack 4 静态分析后做 code-split, **不调到对应代码时这些 chunk 不下载**, runtime bundle 体积无影响. 仅 node_modules 多 80MB.

> 注: webpack 5 / Vite / Parcel 等支持 dynamic import 跳过的现代打包器, **完全不会有这个 warning**, 也不需任何配置. 仅 webpack 4 / 老 vue-cli 4 受影响.

### 完整 Vue 2.6.12 + vue-cli 4 用法

```bash
# 必装
npm i ooxml-excel-editor vue@2.6 @vue/composition-api exceljs
# 按需 (用啥装啥)
npm i echarts jspdf hyperformula
```

```js
// src/main.js
import Vue from 'vue'
import VueCompositionAPI from '@vue/composition-api'
Vue.use(VueCompositionAPI)
```

```vue
<!-- 你的页面 -->
<template>
  <ExcelViewer :src="src" style="height: 100vh" @rendered="onRendered" />
</template>

<script>
import ExcelViewer from 'ooxml-excel-editor/vue2'
import 'ooxml-excel-editor/dist/style.css'
import 'ooxml-excel-editor/dist/vue2.css'

export default {
  components: { ExcelViewer },
  data: () => ({ src: null }),
  methods: { onRendered(wb) { console.log('sheets:', wb.sheets.length) } },
}
</script>
```

---

## 7. peer dep 矩阵

```json
"peerDependencies": {
  "vue": "^2.6.0 || ^2.7.0 || ^3.4.0",
  "@vue/composition-api": "^1.7.0",
  "react": "^17.0.0 || ^18.0.0 || ^19.0.0",
  "react-dom": "^17.0.0 || ^18.0.0 || ^19.0.0",
  "exceljs": "^4.4.0"
}
```

Vue 2 用户(2.6 或 2.7) 装两个: `vue@2.6` 或 `vue@2.7` + `@vue/composition-api`. **Vue 2.6 用户额外**需要在 main.js 调 `Vue.use(VueCompositionAPI)`. Vue 2.7+ 不需要 Vue.use (plugin 自检测 + noop).

React / Vue 3 用户不装这些, 不影响。

**按需依赖** (1.3.2 起从 peer 移除, 改文档说明 — 避免 npm 7+ ERESOLVE 冲突):

```bash
npm i echarts        # 用图表 (动态 import)
npm i jspdf          # 导出 PDF (动态 import)
npm i hyperformula   # 编辑模式 + 公式重算 (动态 import)
```

未装时代码 `await import()` 失败抛友好错误, 不影响其他功能.

---

## 8. 参考实现

- 主壳: [src/vue2/ExcelViewer.ts](../src/vue2/ExcelViewer.ts) — render function 实现, ~600 行
- 加载 composable: [src/vue2/use-excel-document.ts](../src/vue2/use-excel-document.ts) — 跟 Vue 3 / React 壳同逻辑
- 入口: [src/vue2/index.ts](../src/vue2/index.ts) — re-export ExcelViewer + 公共类型
- CSS: [src/vue2/excel-viewer.css](../src/vue2/excel-viewer.css) — 跟 Vue 3 style.css 互补 (`.ov-*` 命名空间)
- demo: [vue2-demo/](../vue2-demo/) — 完整可运行示例

---

## 9. 升级到 1.3.0 正式版前的 alpha 提示

当前 `1.3.0-alpha.X` 是预发版, 用 `--tag alpha` 发布:

```bash
npm install ooxml-excel-editor@alpha vue@2.7
```

正式 1.3.0 发布前可能调整:
- 包体积优化 (共享 chunk)
- 类型声明 (`dist/vue2.d.ts`)
- Vue 2 e2e 覆盖
- 插槽 + 自定义工具栏配置 prop
