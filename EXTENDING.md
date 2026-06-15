# 二开 / 扩展 API 手册(EXTENDING)

> 给**二开者**(在不改源码的前提下定制外观/行为/UI,或写插件)。**纯调用方**(装好、用 props/事件/编辑/导出)看 [README](README.md) 即可。
> 想了解内部结构、在哪改代码:见 [ARCHITECTURE.md](ARCHITECTURE.md)。
>
> 这些扩展点 API 从任一入口都能 import(`ooxml-excel-editor` / `/react` / `/vue2` / `/core` **四入口同源**)。

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
`load(src)` / `getWorkbook()` / `getActiveSheet()` / `setActiveSheet(i)` / `getSelection()` / `setSelection(range)` / `scrollToCell(row,col,{select?})` / `rectOf(row,col)` / `rectOfRange(range)` / `redraw()`,以及下面的导出方法;**编辑命令式 API**(`editCell`/`setStyle`/`createPivotTable`/`openPivotTableDialog`/`createPivotTableFromSelection`/`getConditionalRules`/`addConditionalRule`/`openConditionalFormatDialog`/`setSelectionNumberFormat`/`openNumberFormatDialog`/`getCellComment`/`setCellComment`/`openCommentEditor`/`replaceCurrent`/`replaceAll`/`insertRows`/`undo`/`exportXlsx`…)见 [编辑](#编辑可选默认只读)。

```ts
// 需组件开启 :pivot-table="true"(默认关闭)+ :editable="true"
viewer.value?.createPivotTable({
  sourceRange: { top: 0, left: 0, bottom: 20, right: 4 },
  output: { kind: 'new-sheet' },
  layout: {
    rows: [0],
    columns: [1],
    filters: [{ field: 2, mode: 'equals', value: '华东' }],
    values: [{ field: 3, summary: 'sum' }],
  },
})
```

### 导出 / 打印
内置工具栏右侧有「导出 ▾」菜单(PNG / PDF / 打印 / **导出设置…**)。「导出设置…」打开对话框,可选**范围**(当前选区 / 当前表 / 全部表)、清晰度、是否含行列号/网格线、纸张方向。也可命令式调用(模板 ref / 插件 `viewer`):

| 方法 | 说明 |
|---|---|
| `exportImage(opts?)` | → `Promise<Blob>`,当前/指定表渲染为图片(png/jpeg/webp) |
| `downloadImage(opts?)` | 导出图片并触发下载 |
| `exportPdf(opts?)` | → `Promise<Blob>`,分页 PDF(需可选依赖 `jspdf`) |
| `downloadPdf(opts?)` | 导出 PDF 并触发下载 |
| `print(opts?)` | 打开系统打印对话框(可另存为 PDF,零依赖) |
| `exportXlsx(opts?)` / `downloadXlsx(opts?)` | → `Promise<Blob>` / 下载 **.xlsx**(默认从模型重建;`{fidelity:'overlay'}` 重载原件叠加,保真更高;需可选依赖 `exceljs`) |
| `exportJson(opts?)` / `downloadJson(opts?)` | → `string` / 下载 **.json**(各表首行作 key 的对象数组,raw 类型值) |
| `exportCsv(opts?)` / `downloadCsv(opts?)` | → `string` / 下载 **.csv**(格式化显示值,带 UTF-8 BOM;`opts.target` 指定表,默认当前表) |

**编辑后导出(.xlsx / JSON / CSV)** —— 三种格式都建在**同一份内存数据层**上(`WorkbookModel`:读 `data-access` + 写 `mutations`),无需为每种格式各写一遍解析,故与渲染所见、彼此之间天然一致。JSON 默认输出 raw 类型值(`{format:true}` 可改显示串);CSV 默认输出格式化显示值(WYSIWYG)。

**.xlsx 两种保真模式**:

- **`rebuild`(默认)** —— **从编辑后模型完整重建**:遍历 cells/公式/样式(字体/填充/边框/对齐/数字格式)/合并/行高列宽/冻结/图片/**条件格式**(1.9.0 起)/**批注**(1.11.0 起) 重组成 ExcelJS 工作簿。干净、所见即所得,但**丢失**原件里我们不建模的部分(数据验证、VBA 宏、工作表保护、复杂 DrawingML/图表 等)。图片导出区分 oneCell/twoCell 锚点 + 子格 EMU 偏移。
- **`overlay`(`exportXlsx({ fidelity: 'overlay' })`)** —— **重载原始 .xlsx,只把编辑后的 值/样式/合并/行高列宽/冻结 叠加上去**,**保留** ExcelJS 能往返的其余部分(条件格式 / 数据验证 / 打印设置 / 定义名 / 图表 等)。组件加载时自动留存原件字节供其使用;缺原件时自动回退 `rebuild`。注:overlay 不反映**增删行列 / 图片**编辑(那类用 `rebuild`)。

公共选项:`target`(`'active'`(默认)/`'all'`/索引/索引数组)、`range`(限定单元格区域)、`scale`(清晰度,默认 2)、`includeHeaders`、`gridlines`、`background`;PDF/打印另有 `format`(a4/a3/letter/`[宽,高]mm`)、`orientation`、`margin`(mm)、`fitToWidth`。

**长任务进度 + 取消 + 内置遮罩(P1 + P1.5)** —— **两层**叠加,默认开箱即用,可逐层覆盖:

#### ① Core 层(协议):`onProgress` + `signal`
所有导出方法(PNG / PDF / XLSX / Print)+ 选区图片批量互转 统一接:
```ts
const ctrl = new AbortController()
try {
  await viewer.value.downloadPdf({
    target: 'all',
    onProgress: (p) => console.log(p.stage, p.ratio, p.label), // 'render'/'compose'/'paginate'/'write'/'zip'/'convert'
    signal: ctrl.signal,
  })
} catch (e) {
  if ((e as Error).name === 'AbortError') console.log('用户取消')
  else throw e
}
ctrl.abort() // 任意时刻取消
```
导出全链路在调度点 `await yieldToEvent()` 让出 UI(防假死)+ 调度前 `checkAborted(signal)`(立刻中断)。`ExcelJS.writeBuffer` / `jsPDF` 内部仍是黑盒(`zip`/`write` 阶段),那两段无法细分,但全程都有可视进度。

#### ② Shell 层(UI):**内置居中模态**(默认开)
**不传任何参数**调 `viewer.downloadPdf()` 等异步方法,壳自动建 `AbortController` + 接 `onProgress` → 显示**居中模态**(stage 标签 + 进度条 + 取消按钮)。用户传入的 `onProgress`/`signal` **仍正常链回调**(并存,不冲突)。

#### ③ 关闭 / 覆盖
| 需求 | 做法 |
|---|---|
| 完全关掉内置遮罩(纯回调) | `<ExcelViewer :export-progress="false">`(Vue)/ `exportProgress={false}` (React) |
| 自渲染(用 Element Plus / Ant Design 等自家组件) | **Vue**:`<template #export-progress="{ state, busy, cancel }">…</template>` 插槽;**React**:`renderExportProgress={({state,busy,cancel}) => <YourModal …/>}` |
| 既要内置又自动注入跟踪 | 默认行为已是 —— 用户传 `{ onProgress, signal }` 仍被链回调,内置 UI 也照常显示 |

覆盖矩阵:**导出**(PDF / PNG / XLSX / Print)、**选区图片批量互转**(P2,壳侧 1.2.0 起返 `Promise<number>` 以接遮罩)。**不包含**:文件解析(parsing 有独立的顶栏进度条,与本遮罩分开)、`copySelection` / `setStyle` 等瞬时操作、模板样式 overlay(P3 重设计后是同步纯函数,耗时可忽略)。

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

### 右键菜单(Plan C:三层开放)

默认 `editable` 时显示内置菜单(复制/粘贴/插入/删除/合并/拆分/自动换行/清除内容 + WPS 图片互转)。三种覆盖方式可同时使用:

**① 加 / 减 / 重排内置项** —— 用 `:contextMenu` 传 transform:
```vue
<ExcelViewer
  :context-menu="(ctx, items) => [
    ...items,
    { separator: true },
    { label: `导出此格 PDF (${ctx.activeCell.row + 1},${ctx.activeCell.col + 1})`, action: () => viewer.downloadPdf() },
  ]"
/>
```
- `ctx`: `{ range, single, activeCell, sheet, workbook, editable }` — 当前选区 + 活动格 + 模型句柄
- `items`: 内置 `MenuItem[]`(`{label, action, disabled, separator}`)—— 加、过滤、重排,返回新数组生效;返 `undefined` / `void` 用原样

**② 接管渲染**(用自家 UI 框架的菜单,如 Element Plus / Radix / Headless UI):
```vue
<ExcelViewer
  :context-menu="false"            <!-- 关闭内置弹层(事件仍触发) -->
  @before-context-menu="(p) => p.preventDefault()"
  @context-menu="(p) => myMenu.show(p.x, p.y, p.items)"
/>
```
- `@before-context-menu` 在内置弹出前触发;调 `payload.preventDefault()` 取消内置;`:contextMenu="false"` 等价于自动 preventDefault
- `@context-menu` 在内置弹出后(或被 prevent 后)触发,拿到 `{ x, y, ctx, items }` —— **总会触发**,自渲染只需监听这个事件
- React:`onBeforeContextMenu` / `onContextMenuShow` 接同形 payload

**③ 命令式打开 / 关闭**(键盘 Shift+F10、工具栏触发、跨层调用):
```ts
viewer.openContextMenu(clickX, clickY)                   // 按当前选区算内置 items
viewer.openContextMenu(clickX, clickY, customItems)       // 直接喂自定义 items
viewer.closeContextMenu()
```

**插件贡献**:`definePlugin({ contextMenu: (ctx, items) => [...] })` —— 多插件按数组顺序串行(后者拿前者输出),组件 `:contextMenu` prop 最后覆盖,顺序固定 `内置 → 插件 → prop`。

`MenuItem` / `ContextMenuCtx` / `ContextMenuTransform` 全部从 `ooxml-excel-editor/core` 导出(TS 类型完整)。

### 操作工具栏(可配置 / 可插件 / 响应式)
顶栏(文件名/导出/缩放)下方有一行**操作工具栏**,内置 `find`/`filter` 按钮默认显示。用 `:toolbar` 配置:
```vue
<ExcelViewer :src="file" />                                           <!-- 默认: find + filter + sort -->
<ExcelViewer :toolbar="['find','filter','separator','zoom','export']" /> <!-- 控制项/顺序/分隔 -->
<ExcelViewer :toolbar="false" />                                      <!-- 隐藏整条 -->
```
- **内置 id**:`find`(查找)、`filter`(切换自动筛选 —— 文件没设也能点出下拉)、`sort`(按活动单元格所在列升序/降序;未开启自动筛选时会先按选区/已用区建立范围)、`clear-filter`(清除筛选,无筛选时禁用)、`copy`(复制选区)、`pivot-table`(透视表入口:选中带表头数据区后选择生成位置,可输出到现有工作表单元格或新建工作表;创建后打开 WPS 风格右侧字段面板,需 `pivotTable` + `editable`,功能未开启时不渲染)、`conditional-format`(条件格式管理入口:列出当前表规则可删/可编辑 + 新建全 6 类规则,需 `conditionalFormat` + `editable`,功能未开启时不渲染)、`number-format`(数字格式编辑入口:分类 + 预览 + 自定义格式代码,需 `editable`)、`format-painter`(格式刷:采样源格样式刷到目标,需 `editable`)、`wrap-text`(自动换行 toggle,WPS 风格,需 `editable`)、`image-tools`(图片工具 ▾:选区/整表/整列 浮动 ⇄ 嵌入互转,需 `editable`)、`template`(模板 ▾:仅 JSON / 模型数据源下生效;导入 .xlsx 当样式捐赠者;xlsx 数据源下禁用)、`freeze`(冻结/取消)、`zoom`(缩放下拉)、`export`(导出/打印下拉)、`'separator'`/`'|'`(分隔线)。
- **富项类型**(`ToolbarItem`):`type:'separator'` 分隔线;`items: ToolbarItem[]` 变下拉子菜单;`disabled?(viewer)` 禁用态;`iconSvg`(内联 SVG,优先于 `icon` emoji)/ `icon` / `label` / `title` / `onClick(viewer)` / `active?(viewer)`。
- **响应式溢出**:宽度不足时,放不下的项自动折叠进「⋯ 更多」下拉。
- **插件贡献**:`ExcelPlugin.toolbar: ToolbarItem[]`,插件加载即追加(opt-in)。
- 内置图标用极简线性 **SVG**(跨平台一致);`filter` 按钮让筛选**看得见**,不必依赖文件自带 autofilter。

### 分层 UI(slots)
具名 slot:`header`(顶栏)/ `toolbar`(作用域 `{ items }`,替换整条操作栏)/ `statusbar` / `loading` / `error` / `empty`(缺省用内置)。
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
import { definePlugin } from 'ooxml-excel-editor'

const highlightNegatives = definePlugin({
  name: 'highlight-negatives',
  theme: { selBorder: '#e91e63' },
  cellStyle: (c) => (typeof c.raw === 'number' && c.raw < 0 ? { font: { color: '#d00' } } : undefined),
  events: { 'cell-click': (p) => console.log('clicked', p) },
  overlay: ({ rectOf }) => {
    const r = rectOf(0, 0)
    if (!r) return null
    const el = document.createElement('div') // 返回 DOM(框架无关,Vue/React 通用)
    el.textContent = '⚑'
    Object.assign(el.style, { position: 'absolute', left: r.x + 'px', top: r.y + 'px' })
    return el
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
插件字段:`theme` / `transformModel` / `cellStyle` / `events`(事件→处理器) / `overlay`(返回 **DOM 节点**,随滚动跟随) / `toolbar`(贡献操作栏按钮 `ToolbarItem[]`) / `setup(ctx)`(拿 `viewer` 命令式 API、`on()` 订阅事件,返回可选清理函数)。

> **跨框架**:插件全字段框架无关,**同一份 `definePlugin` 在 Vue 和 React 壳通用**(`overlay` 返回 DOM 而非 VNode)。React 用法:`<ExcelViewer plugins={[myPlugin]} />`。


## Headless / Node 安全 API 面

`ooxml-excel-editor/core` 的下列出口**零浏览器依赖,可在纯 Node(无 DOM/canvas)直接用**。用法与可跑示例见 [README → Node 用法](./README.md#node--服务端-headless-用法) 与 [`examples/`](./examples)。

**可用(纯 Node)**:
- 打开 / 解析:`openWorkbook(src)`(一行门面)、`parseWorkbook(buffer)`(收 `ArrayBuffer | Uint8Array`,Node `Buffer` 直接传)、`loadArrayBuffer(src)`、`jsonToWorkbook(data, opts)` / `makeDefaultStyle()`(数据直建模型)。
- 取数:`getCellText` / `getCellValue` / `getSheetData` / `getRangeData` / `sheetToJSON` / `getWorkbookJSON` / `cellDisplayText` / `formatValue`。
- 编辑模型(框架无关,不经 viewer):`setCellValue` / `clearCell` / `setRangeValues` / `applyStyleOverride` / `mergeStyleOverride` / `setColumnWidth` / `setRowHeight` / `insertRows` / `deleteRows` / `insertCols` / `deleteCols` / `addImage` / `removeImage`,以及 `cloneWorkbook` / 命令栈 `applyCommand`。
- 公式重算:内置 `builtinFormulaEngineFactory`(MIT、零依赖,默认)或注入 `hyperFormulaEngineFactory`(需装 `hyperformula`)。
- 导出:`workbookToXlsxBytes(wb, opts)` → `Uint8Array`(`fs.writeFileSync` 直接落盘;`fidelity: 'overlay' + sourceBuffer` 保真往返)、`toCsv` / `toWorkbookJson`。

**不可用(硬依赖浏览器 canvas / DOM,Node 调用即抛错)**:
- `workbookToXlsxBlob`(返回 `Blob` —— Node 改用 `workbookToXlsxBytes`)、`canvasToBlob` / `canvasToDataURL` / `downloadBlob`。
- 图片 / PNG / JPEG / **PDF 导出**、`print()`、`CanvasRenderer` / `ViewerController` 渲染、内置 `DefaultEditor`。
- 注:`finalizeImages` 在 Node 安全跳过(图片保留 `bytes`/`mime`,不生成 blob URL);`loadArrayBuffer` 的 URL 字符串分支走 `fetch`(不支持本地路径 / `file://`)—— Node 用 `fs` 读 Buffer 传入。
