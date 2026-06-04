# ooxml-excel-preview

> Vue 3 高保真 **.xlsx 预览组件** —— Canvas 渲染,只读,无编辑。从零实现解析与渲染,尽量还原微软 Excel 打开工作簿的观感。

[English](#english) · 中文

## 特性

- 📊 **Canvas 高保真渲染**:DPR 高清、虚拟滚动(万行流畅)、冻结窗格四象限
- 🔢 **自写数字格式引擎**:千分位/货币/百分比/科学计数/分数、四段格式(正;负;零;文本)、`[Red]` 颜色、`[>=100]` 条件段、中文日期 `yyyy"年"`、`[h]:mm` 经过时间
- 🗓 **日期序列号**:含 Excel 1900 闰年 bug、1904 系统
- 🎨 **主题色 + tint**、indexed 调色板、合并单元格、边框(细/粗/虚/双线)、填充(纯色/图案/渐变)
- 🌈 **条件格式**:色阶 / 数据条 / 图标集 / cellIs / top10
- 🖼 **图片 + 图表**(DrawingML → ECharts 近似还原)、**形状/文本框**、**迷你图**(sparklines)、**批注**、**数据验证**下拉、**自动筛选**样式
- 📝 **文本溢出**到相邻空格、**自动行高**
- 🖱 **交互**:单元格选区(合并感知)、拖选、公式栏、状态栏(计数/求和/均值/最值)、超链接可点、裁切文本悬停看全文、Ctrl+C 复制为 TSV、**Ctrl+F 查找**(高亮 + 上/下定位 + 计数 + 区分大小写/全字匹配)
- 🖨 **导出 / 打印**:整表/选区/多表导出 **PNG/JPEG**、**PDF**(位图 + **矢量·文字可选可搜**两种)、**系统打印**(可另存 PDF);默认还原原生 `pageSetup`(纸张/方向/页边距/缩放/打印区域/**打印标题行列每页重复**);宽表**横向跨页**(页矩阵);`beforeRenderPage` 注入页眉/页脚/水印、`configureDoc` 注册字体;内置「导出设置」对话框
- ⚡ **按需加载**(无图表文件不下载 echarts、不导出 PDF 不下载 jspdf)、**友好错误兜底**(损坏/加密/旧 .xls)、解析失败自动给出可读提示

> 预览不需要公式引擎 —— .xlsx 缓存了公式结果,直接显示。详见 [EXCEL还原难点.md](./EXCEL还原难点.md)。

## 安装

```bash
npm i ooxml-excel-preview vue exceljs
# echarts 可选:仅当要渲染图表时才需要
npm i echarts
# jspdf 可选:仅当要导出 PDF 时才需要(打印/图片导出不需要)
npm i jspdf
```

`vue` / `exceljs` 是必需 **peerDependencies**;`echarts` / `jspdf` 为**可选** peer —— 未安装时分别只影响"图表渲染""PDF 导出",其余功能正常,且不会被打包进你的产物(运行时才动态加载)。

## 使用

### 组件方式

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { ExcelViewer } from 'ooxml-excel-preview'
import 'ooxml-excel-preview/style.css'

const file = ref<File>()
</script>

<template>
  <ExcelViewer
    :src="file"
    :file-name="file?.name"
    style="height: 100vh"
    @rendered="(wb) => console.log('已渲染', wb.sheets.length, '个工作表')"
    @error="(msg) => console.error(msg)"
  />
</template>
```

### 全局注册(Vue 插件)

```ts
import OoxmlExcelPreview from 'ooxml-excel-preview'
import 'ooxml-excel-preview/style.css'

app.use(OoxmlExcelPreview) // 注册全局组件 <ExcelViewer />
```

### 程序化解析(只要数据模型,不渲染)

```ts
import { parseWorkbook, loadArrayBuffer } from 'ooxml-excel-preview'

const buffer = await loadArrayBuffer(file) // File/Blob/ArrayBuffer/Uint8Array/URL
const wb = await parseWorkbook(buffer)
console.log(wb.sheets[0].cells)
```

## API

### `<ExcelViewer>`

| Prop | 类型 | 说明 |
|---|---|---|
| `src` | `File \| Blob \| ArrayBuffer \| Uint8Array \| string(URL)` | 要预览的 .xlsx 数据源 |
| `fileName` | `string` | 标题栏显示的文件名(可选) |

| 事件 | 载荷 | 触发时机 |
|---|---|---|
| `rendered` | `WorkbookModel` | 解析并首次渲染完成 |
| `error` | `string` | 解析失败(友好文案) |

容器需要有明确高度(组件填满父容器),例如 `style="height: 100vh"`。

### 具名导出

| 导出 | 说明 |
|---|---|
| `ExcelViewer` | 预览组件 |
| `parseWorkbook(buffer)` | `ArrayBuffer → Promise<WorkbookModel>`(优先 Web Worker) |
| `loadArrayBuffer(src)` | 多种输入归一化为 `ArrayBuffer` |
| `default` | Vue 插件(`app.use`) |
| 类型 | `WorkbookModel` / `SheetModel` / `CellModel` / `CellStyle` / `MergeRange` / `ConditionalRule` / `ChartSpec` / `ImageAnchor` / `CssColor` / `ExcelSource` |

## 扩展 API(不改源码定制)

组件按"分层可扩展"设计 —— 用内置 props/events/slots/命令式 API 即可定制外观、行为、数据,并在网格上叠自己的 UI。

### 外观主题 `:theme`
```vue
<ExcelViewer :src="file" :theme="{ gridLine: '#e8e8e8', selBorder: '#e91e63', selFill: 'rgba(233,30,99,.1)' }" />
```
可覆盖:`headerBg / headerText / headerLine / gridLine / selBorder / selFill`(见 `ViewerTheme` / `DEFAULT_THEME` 导出)。

### 数据 / 渲染钩子
```vue
<ExcelViewer
  :src="file"
  :transform-model="(wb) => { wb.sheets[0].name = '改过的名字'; return wb }"
  :cell-style="(cell) => typeof cell.raw === 'number' && cell.raw < 0 ? { font: { color: '#d00' } } : undefined"
/>
```
- `transformModel(wb)`:解析后、渲染前改模型(返回新模型或就地改)。
- `cellStyle(cell, {row,col})`:按条件覆盖单元格样式(`font/fill/borders` 浅合并)。

### 事件
| 事件 | 载荷 |
|---|---|
| `cell-click` / `cell-dblclick` | `{ row, col, text }` |
| `selection-change` | `{ range, active }` |
| `sheet-change` | `{ index, name }` |
| `hyperlink-click` | `{ url, cell }`(配 `:open-links="false"` 接管跳转) |
| `rendered` / `error` / `progress` | 见上 |

### 命令式 API(模板 ref)
`load(src)` / `getWorkbook()` / `getActiveSheet()` / `setActiveSheet(i)` / `getSelection()` / `setSelection(range)` / `rectOf(row,col)` / `rectOfRange(range)` / `redraw()`,以及下面的导出方法。

### 导出 / 打印
内置工具栏右侧有「导出 ▾」菜单(PNG / PDF / 打印 / **导出设置…**)。「导出设置…」打开对话框,可选**范围**(当前选区 / 当前表 / 全部表)、清晰度、是否含行列号/网格线、纸张方向。也可命令式调用(模板 ref / 插件 `viewer`):

| 方法 | 说明 |
|---|---|
| `exportImage(opts?)` | → `Promise<Blob>`,当前/指定表渲染为图片(png/jpeg/webp) |
| `downloadImage(opts?)` | 导出图片并触发下载 |
| `exportPdf(opts?)` | → `Promise<Blob>`,分页 PDF(需可选依赖 `jspdf`) |
| `downloadPdf(opts?)` | 导出 PDF 并触发下载 |
| `print(opts?)` | 打开系统打印对话框(可另存为 PDF,零依赖) |

公共选项:`target`(`'active'`(默认)/`'all'`/索引/索引数组)、`range`(限定单元格区域)、`scale`(清晰度,默认 2)、`includeHeaders`、`gridlines`、`background`;PDF/打印另有 `format`(a4/a3/letter/`[宽,高]mm`)、`orientation`、`margin`(mm)、`fitToWidth`。

**默认还原 OOXML 原生页面设置** —— PDF/打印时,未显式指定的 `format`/`orientation`/`margin`/`fitToWidth` 自动取自工作表的 `pageSetup`(纸张、方向、页边距、适应页面/缩放),并应用**打印区域**(默认导出范围)与**打印标题行/列**(每页顶部/左侧重复)。显式传入的选项始终覆盖之。

**分页** —— `fitToWidth: true`(默认)把整表缩放到页宽、只竖向跨页;`fitToWidth: false`(或工作表未设"适应页面")按自然尺寸×缩放,**宽表横向跨页 + 高表竖向跨页**(像 Excel 的页矩阵,顺序"先下后右"),此时打印标题列在每张横向页左侧重复。

**`beforeRenderPage` 扩展钩子** —— 每页贴图后调用,拿到 `jsPDF` 实例画页眉/页脚/水印/页码:
```ts
const viewer = ref()  // <ExcelViewer ref="viewer" />
await viewer.value.downloadPdf({
  target: 'all',
  beforeRenderPage: ({ doc, pageIndex, pageCount, pageWidth, pageHeight, margin, sheetName }) => {
    doc.setFontSize(9); doc.setTextColor(120)
    doc.text(sheetName, margin.left, pageHeight - 5)
    doc.text(`第 ${pageIndex + 1} / ${pageCount} 页`, pageWidth - margin.right, pageHeight - 5, { align: 'right' })
    doc.setFontSize(56); doc.setTextColor(230)
    doc.text('PREVIEW', pageWidth / 2, pageHeight / 2, { align: 'center', angle: 30 })  // 水印
  },
})
```
打印另有 `title` / `headerHtml` / `footerHtml`(每页 HTML 片段)。
> 图片/图表/形状是 DOM 叠加层,导出时会自动合成到底图;"导出全部表"中非当前表的图表需 `echarts` 可用才能离屏渲染。

#### 矢量 PDF(文字可选可搜)

两种 PDF 并存,工具栏菜单有「位图 / 矢量」两项,API 用 `vector` 切换:
```ts
await viewer.value.downloadPdf({ vector: true })
```
- **位图 PDF**(默认):整表贴图,完整还原观感。
- **矢量 PDF**:逐格用真文字 + 矢量填充/边框绘制 —— 文字**可选中、可搜索、放大清晰、文件更小**。条件格式(背景色/数据条/图标)也走矢量绘制;仅**迷你图、旋转文字、富文本**这几类格会自动从底图**裁小图兜底**(内容不丢)。

**中文字体** —— jsPDF 内置字体只认拉丁/数字。用 `configureDoc(doc)` 钩子注册中文 TTF 即可全矢量;不注册时,含中文的单元格自动转为该格小图(清晰但不可选):
```ts
await viewer.value.downloadPdf({
  vector: true,
  configureDoc: (doc) => {
    doc.addFileToVFS('NotoSansSC.ttf', base64Ttf)  // 你的中文字体(建议子集化)
    doc.addFont('NotoSansSC.ttf', 'NotoSC', 'normal')
    doc.setFont('NotoSC')                            // 设为默认 → 中文也走矢量
  },
})
```
> 提示:中文表格若不注册字体,矢量模式会产生很多小图、文件偏大且较慢 —— 注册一个子集字体即可全矢量。

### 分层 UI(slots)
具名 slot:`toolbar` / `statusbar` / `loading` / `error` / `empty`(缺省用内置)。
**作用域 `overlay` slot** —— 在格子上叠自己的 Vue 组件,随滚动/缩放跟随:
```vue
<ExcelViewer :src="file">
  <template #overlay="{ rectOf, tick }">
    <!-- tick 变化触发重算;rectOf(row,col) 给当前屏幕矩形 -->
    <button v-if="rectOf(2,1)" :style="posStyle(rectOf(2,1), tick)" @click="...">★</button>
  </template>
</ExcelViewer>
```
覆盖层容器 `pointer-events:none`(滚动穿透),子元素自动 `pointer-events:auto`(可点)。

### 插件 `definePlugin`
把上面所有扩展点(主题/数据钩子/渲染钩子/事件/overlay/命令式 API)打包成一个插件,`:plugins` 分发;多个插件按数组顺序合并,组件自身 props 最后覆盖。
```ts
import { definePlugin } from 'ooxml-excel-preview'

const highlightNegatives = definePlugin({
  name: 'highlight-negatives',
  theme: { selBorder: '#e91e63' },
  cellStyle: (c) => (typeof c.raw === 'number' && c.raw < 0 ? { font: { color: '#d00' } } : undefined),
  events: { 'cell-click': (p) => console.log('clicked', p) },
  overlay: ({ rectOf }) => {
    const r = rectOf(0, 0)
    return r ? h('div', { style: { position: 'absolute', left: r.x + 'px', top: r.y + 'px' } }, '⚑') : null
  },
  setup: ({ viewer, on }) => {
    on('selection-change', (s) => console.log(s))
    // viewer.setSelection(...) / viewer.getWorkbook() ...
    return () => {/* 清理 */}
  },
})
```
```vue
<ExcelViewer :src="file" :plugins="[highlightNegatives]" />
```
插件字段:`theme` / `transformModel` / `cellStyle` / `events`(事件→处理器) / `overlay`(返回 VNode,随滚动跟随) / `setup(ctx)`(拿 `viewer` 命令式 API、`on()` 订阅事件,返回可选清理函数)。

## 浏览器支持

现代浏览器(Chrome/Edge 80+、Safari 15+、Firefox 114+,需支持 Canvas / ResizeObserver)。

> **解析线程**:发布的组件库在**主线程**解析(`exceljs` 为 peer 依赖,不重复打包)。本仓库的 demo/dev 额外启用了 **Web Worker** 解析(大文件不卡 UI)。如果你的应用要处理很大的文件,可直接用导出的 `parseWorkbook` 包进你自己的 Worker。

## 范围边界(第一版不做)

- 编辑 / 公式重算(只显示缓存结果)
- 透视表:**数据按普通单元格显示**,但无字段按钮/下拉等透视专属 UI
- SmartArt;形状仅支持 rect/roundRect/ellipse + 文本(复杂自定义几何按矩形近似)
- `.xls`(旧 BIFF 二进制)/ 加密文件(给出友好提示)
- 图表为 ECharts 近似,非像素级一致
- 导出为**位图**(PNG/PDF 内嵌图片),非矢量;PDF 按页宽缩放后竖向分页

## 开发

```bash
npm install
npm run dev            # 本地预览(demo)
node scripts/gen-sample.mjs   # 生成 public/sample.xlsx 示例
npm run test           # 单元测试(node 环境,纯逻辑)
npm run test:e2e       # 真浏览器 e2e(Playwright):canvas 渲染 + jsPDF 导出 + 下载全链路
npm run typecheck      # 类型检查
npm run build          # 构建组件库(dist/)
npm run build:demo     # 构建 demo 站点
```

> **e2e 说明**:`npm run test:e2e` 用 Playwright 起 dev 服务 + 无头 Chromium,加载示例 → 渲染 → 导出 PNG/位图PDF/矢量PDF,校验产物(PNG 魔数、`%PDF`、矢量 PDF 的文字操作符数量多于位图)。覆盖 node 单测做不到的真实 canvas/jsPDF 绘制。首次需 `npx playwright install chromium` 下载浏览器(本仓库 `@playwright/test` 固定 `1.58.0` 对应 chromium-1208)。

## License

MIT

---

<a name="english"></a>
## English

A **Vue 3 high-fidelity, read-only `.xlsx` preview component** with a from-scratch parser and canvas renderer. Renders cells, number formats, merges, conditional formatting, images, charts (via ECharts), sparklines, comments, data validation, frozen panes, and supports selection / copy / hyperlinks. Parsing runs in a Web Worker (with main-thread fallback). `vue` / `exceljs` are peer dependencies; `echarts` is an optional peer (only needed for charts).

```bash
npm i ooxml-excel-preview vue exceljs
```

```ts
import { ExcelViewer } from 'ooxml-excel-preview'
import 'ooxml-excel-preview/style.css'
```

See the API table above. MIT licensed.
