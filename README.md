# ooxml-excel-editor

[![npm version](https://img.shields.io/npm/v/ooxml-excel-editor.svg)](https://www.npmjs.com/package/ooxml-excel-editor)
[![npm downloads](https://img.shields.io/npm/dm/ooxml-excel-editor.svg)](https://www.npmjs.com/package/ooxml-excel-editor)
[![CI](https://github.com/ojaDeveloper/ooxml-excel-editor/actions/workflows/ci.yml/badge.svg)](https://github.com/ojaDeveloper/ooxml-excel-editor/actions/workflows/ci.yml)
[![license](https://img.shields.io/npm/l/ooxml-excel-editor.svg)](./LICENSE)

> Vue 3 + **Vue 2** + React 高保真 **.xlsx 预览 / 编辑组件** —— Canvas 渲染,**默认只读预览,可选开启编辑**。从零实现解析与渲染,尽量还原微软 Excel 打开工作簿的观感。**三个壳 UI 1:1 对齐**(Vue 3 SFC 是标准,Vue 2 / React 复刻)。

**🔗 在线 demo(直接试用):https://ojadeveloper.github.io/ooxml-excel-editor/** · [English](#english) · 中文

## ⚡ 快速开始

**装**(按框架二选一):

```bash
npm i ooxml-excel-editor vue                          # Vue 3 (默认入口)
npm i ooxml-excel-editor react react-dom              # React 壳 (/react 子入口)
npm i ooxml-excel-editor vue@2.7 @vue/composition-api # Vue 2.6/2.7+ (/vue2 子入口, 1.3.0+)
```

> 1.3.2+ `exceljs` / `jspdf` / `hyperformula` 已 **inline 编进 dist chunks/** (ES2017 降级 + 老打包器零解析), `echarts` 仍 external 走 `dependencies` (npm 自动装, 避免主题 dual instance). 消费方仅装 framework (`vue` / `react` / `@vue/composition-api`), 其他 lib 完全不用手动装. `dist` ~5.4 MB, tgz ~1.33 MB.

**用**(Vue,容器要给高度;`src` 可传 `File` / `Blob` / `ArrayBuffer` / `Uint8Array` / URL 字符串):

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { ExcelViewer } from 'ooxml-excel-editor'
import 'ooxml-excel-editor/style.css'
const src = ref<File>() // 绑个 <input type="file" @change> 给它即可
</script>

<template>
  <ExcelViewer :src="src" style="height: 100vh" />
</template>
```

默认**只读预览**;想编辑加 `:editable="true"`(React / Vue 2 同名 `editable`)。React / Vue 2 写法、props/事件表、编辑 / 导出 API 见下文对应章节。Vue 2 用法详见 [docs/Vue2.md](./docs/Vue2.md)。

> 纯使用者只需读 **安装 / 使用 / API / 编辑 / 导出** 几节即可接入,无需看源码;类型随包发 `.d.ts`(IDE 自动补全)。「扩展 API / 插件 / 开发」是进阶,可跳过。

## 特性

- 📊 **Canvas 高保真渲染**:DPR 高清、虚拟滚动(万行流畅)、冻结窗格四象限、**滚动自动延伸空行/列**(像 Excel 无限网格,但不写进 dimension/文件)
- 🔢 **自写数字格式引擎**:千分位/货币/百分比/科学计数/分数、四段格式(正;负;零;文本)、`[Red]` 颜色、`[>=100]` 条件段、中文日期 `yyyy"年"`、`[h]:mm` 经过时间
- 🗓 **日期序列号**:含 Excel 1900 闰年 bug、1904 系统
- 🎨 **主题色 + tint**、indexed 调色板、合并单元格、边框(细/粗/虚/双线)、填充(纯色/图案/渐变)
- 🌈 **条件格式**:色阶 / 数据条 / 图标集 / cellIs / top10 渲染;开 `conditionalFormat` 后可**新建/编辑/删除全 6 类规则**(工具栏入口 + API,导出回写 .xlsx)
- 🖼 **图片 + 图表**(DrawingML → ECharts 近似还原)、**形状/文本框**、**迷你图**(sparklines)、**批注**、**数据验证**(列表型点格内箭头弹选可撤销;整数/小数/日期/文本长度等在编辑模式**拦截非法输入** + WPS 式出错/输入提示)、**自动筛选**样式
- 📌 **WPS 单元格内嵌图(DISPIMG)**:识别并展示 WPS 私有的"嵌在格里的图"(普通工具会缺图);编辑模式下支持**一键浮动 ⇄ 嵌入互转**。见 [WPS 单元格内嵌图](#wps-单元格内嵌图dispimg)
- 🔍 **图片点击放大 + 下载原图**:网格里的图(内嵌图/浮动图)点开看大图、下载原件。只读模式单击图放大、编辑模式右键「查看大图」。`imageLightbox` prop 控制(默认开),`openImageLightbox(src)` 命令式打开。
- 📋 **从 Excel/WPS 富粘贴**:`Ctrl+V` 解析剪贴板 HTML → 还原字体/颜色/填充/边框/对齐/合并单元格,整块单次撤销。**Excel/WPS 把格式放在 `<style>` 块的 CSS 类里(`<td class="xl65">`),解析时会把类规则合并进每格** —— 不只读内联 `style=`。**`Ctrl+V` 走 `paste` 事件拿原始 HTML**;`navigator.clipboard.read()`(右键菜单粘贴用)会**净化**删掉 `<style>`/注释,所以右键粘贴从 WPS 拿的格式不如 `Ctrl+V` 全。图片走多通道:data-uri `<img>` / **WPS VML `o:gfxdata`**(区域复制的内嵌图藏在 VML 里,是个 zip,解出来落格)/ 单图 blob / 拖文件;**数字格式**(日期/货币)也从 `mso-number-format` 解析还原,不再变成裸序列号。**注**:Excel 某些版本只给 `file:///` 本地路径的 `<img>`(浏览器读不到)而不带 `o:gfxdata`,那种区域图仍救不回。
- 📋 **应用内 1:1 复制粘贴**:本组件自己 `Ctrl+C` 的内容会把**完整模型快照**嵌进剪贴板(`<table data-ooxml-clip>`),`Ctrl+V` 时识别并 1:1 还原 —— 数字不会退化成文本、边框/数字格式/合并/DISPIMG 图片/**行高**全保留;因为快照随剪贴板走,**Vue3/Vue2/React 三壳之间、跨标签页互相复制结果都一致**。粘到外部应用(Excel/WPS/Word)则读可见 `<table>`(近似)。**列宽以目标现有表头为准、不被源覆盖**(列宽整列共享,改了会动表头;同 Excel 默认粘贴)。
- 📝 **文本溢出**到相邻空格、**自动行高**
- 🖱 **交互**:单元格选区(合并感知)、拖选、公式栏、状态栏(计数/求和/均值/最值)、超链接可点、裁切文本悬停看全文、Ctrl+C 复制(**同应用内 1:1 保真**:含数字原始值/数字格式/边框/合并/图片/行高,跨 Vue3/Vue2/React 实例互相复制都一致;**列宽以目标表头为准不覆盖**;另带 TSV/HTML 供贴进 Excel/WPS)、**Ctrl+F 查找替换**(高亮 + 上/下定位 + 计数 + 区分大小写/全字匹配;编辑模式带替换 / 全部替换)、**自动筛选**(点下拉真能筛:去重值多选 + 搜值 + 清除)、**自动填充柄**(编辑模式拖选区右下角填序列:等差/日期/星期月份/文本递增,见 [编辑](#编辑可选默认只读))、**不连续多选**(Ctrl/⌘ 点击行头/列头/格 加选不相邻区域,复制堆叠 + 状态栏跨区统计)
- 🖨 **导出 / 打印**:整表/选区/多表导出 **PNG/JPEG**、**PDF**(位图 + **矢量·文字可选可搜**两种)、**系统打印**(可另存 PDF);默认还原原生 `pageSetup`(纸张/方向/页边距/缩放/打印区域/**打印标题行列每页重复**);宽表**横向跨页**(页矩阵);`beforeRenderPage` 注入页眉/页脚/水印、`configureDoc` 注册字体;内置「导出设置」对话框
- ⚡ **按需加载**(无图表文件不下载 echarts、不导出 PDF 不下载 jspdf)、**友好错误兜底**(损坏/加密/旧 .xls)、解析失败自动给出可读提示

- 📤 **数据读取 API**:不必自己再解析 —— `getCellText`/`getSheetData`/`sheetToJSON`/`getRangeData`(独立函数 + 组件 ref 方法),值/显示文本可选,合并/日期/数字格式都处理好
- ✏️ **编辑(可选,默认只读)**:开 `editable` 即可编辑 —— 单元格值 / 样式(粗体/对齐/填充)/ 列宽行高 / 浮动图片(拖拽移改)/ 增删行列;**撤销重做**(Ctrl+Z/Y)、**前后完整快照事件**、**脏状态 + 一键还原原件**;可换**公式引擎**自动重算依赖格;可注入**自定义编辑器**(下拉/日期/图片选择器);**导出回 .xlsx / JSON / CSV**(所见即所得)。见 [编辑](#编辑可选默认只读)

> 纯预览不需要公式引擎 —— .xlsx 缓存了公式结果,直接显示;仅开启**编辑 + 重算**时才用(可选 `hyperformula`)。详见 [EXCEL还原难点.md](./EXCEL还原难点.md)。

## 安装

一个包,**四个子入口** —— 框架无关的 core 引擎被 Vue 3 / React 两个壳共享(`dist/core.js` 只打一份),Vue 2 因 SFC 编译器跟 Vue 3 冲突独立打包(内嵌 core)。**只需按框架装对应 framework**,`exceljs` / `fflate` / `jspdf` / `hyperformula` 等重依赖 **1.3.2+ 已内联进 dist,无需手动装**:

```bash
# Vue 3 项目
npm i ooxml-excel-editor vue

# React 项目
npm i ooxml-excel-editor react react-dom

# Vue 2.6.x 或 2.7+ 项目 (1.3.0+) — 必装 @vue/composition-api (兼容 2.6 + 2.7)
npm i ooxml-excel-editor vue@2.7 @vue/composition-api
# Vue 2.6.x 还需 main.js: Vue.use(require('@vue/composition-api').default)

# 纯 Node / 只解析读数据 / 导出(不渲染 UI,无框架)
npm i ooxml-excel-editor
```

四个入口:

| import | 内容 | 需要装的 framework | 体积 (gzip) |
|---|---|---|---|
| `ooxml-excel-editor` | **Vue 3** 组件 `<ExcelViewer>` (参考实现 Standard) | `vue@3` | ~19 KB + 共享 chunks |
| `ooxml-excel-editor/react` | **React** 组件 `<ExcelViewer>` (1:1 复刻 Vue 3) | `react` + `react-dom` | ~11 KB + 共享 chunks |
| `ooxml-excel-editor/vue2` | **Vue 2.6 / 2.7+** 组件 `<ExcelViewer>` (1:1 复刻 Vue 3) | `vue@2.6+` + `@vue/composition-api` | ~124 KB (内嵌 core) |
| `ooxml-excel-editor/core` | 框架无关引擎(解析/渲染/控制器/导出/读数据) | 无(纯 Node 也可用,见 [Node 用法](#node--服务端-headless-用法)) | ~1 KB + 共享 chunks |

`vue` / `react` / `vue@2` 按框架三选一(均为可选 peer);`exceljs` / `fflate` / `jspdf` / `hyperformula` **已内联进 dist**(无需装、也**绝不重复进你的产物**,运行时才动态从 chunk 加载);`echarts` 是 external 依赖(npm 自动装,仅图表渲染才真正加载,避免主题 dual instance)。

> **三壳 UI 1:1**: Vue 3 SFC 是参考实现 (Standard), Vue 2 / React 1:1 复刻视觉与交互 (工具栏 SVG 图标 / 下拉子菜单 / 公式栏 / 状态栏 / dialog / 浮层 / 演示 demo 全部对齐). 详见 [docs/Vue2.md](./docs/Vue2.md) 跟 Vue 3 的差异速查 + [CLAUDE.md](./CLAUDE.md) 第 7 中心原则。

> 公式重算的引擎(1.14.0 起):默认是**内置 MIT 引擎**(零依赖,覆盖 ~60 常用函数,无许可证负担)。需要更全函数集时,`formulaEngine` prop 注入 **HyperFormula**(`hyperFormulaEngineFactory`,GPL-3.0/商业双授权,~395 函数)或自研引擎(实现 `FormulaEngine` 接口)。

## Node / 服务端 (headless) 用法

不渲染 UI、纯在 **Node**(无浏览器、无 canvas)里处理 .xlsx —— 走 `ooxml-excel-editor/core`(框架无关、ESM)。装包只需 `npm i ooxml-excel-editor`(`exceljs` 等已内联)。可跑示例见 [`examples/`](./examples)。

**适合 Node 的两件事**(比裸 `exceljs` 多了"显示文本渲染 / 合并 / 日期 / 富文本"的保真):

**① 解析取数** —— 拿"人看到的"数据(显示文本 / JSON / CSV):

```ts
import { readFileSync } from 'node:fs'
import { openWorkbook, getSheetData, sheetToJSON, getCellText, toCsv } from 'ooxml-excel-editor/core'

// openWorkbook 直接吃 Node Buffer —— 不必手动转 ArrayBuffer
const wb = await openWorkbook(readFileSync('input.xlsx'))
const sheet = wb.sheets[0]

getCellText(sheet, 0, 0)               // 单格显示文本(数字格式/日期已渲染,如 "2021年1月")
getSheetData(sheet, { format: true })  // 2D 数组(format:false 给原始值)
sheetToJSON(sheet)                      // 对象数组(首行当表头)
toCsv(sheet)                            // CSV 文本
```

**② 高保真往返编辑** —— 打开真实 .xlsx → 程序化改值/样式 → 保样式回写(`overlay` 保留原件的样式/条件格式/图片/透视表,裸 exceljs 会丢):

```ts
import { readFileSync, writeFileSync } from 'node:fs'
import { openWorkbook, setCellValue, applyStyleOverride, workbookToXlsxBytes } from 'ooxml-excel-editor/core'

const src = readFileSync('input.xlsx')
const wb = await openWorkbook(src)
const sheet = wb.sheets[0]

setCellValue(sheet, 1, 2, 123.45)
applyStyleOverride(sheet, 1, 2, { font: { bold: true, color: '#FF0000' } })

// workbookToXlsxBytes 返回 Uint8Array(不是浏览器 Blob),直接 fs 落盘
const bytes = await workbookToXlsxBytes(wb, {
  fidelity: 'overlay',
  sourceBuffer: src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength),
})
writeFileSync('output.xlsx', bytes)
```

**③ 从数据建表**(可选;纯建表裸 `exceljs` 也行)—— `jsonToWorkbook` → `workbookToXlsxBytes`:

```ts
import { writeFileSync } from 'node:fs'
import { jsonToWorkbook, workbookToXlsxBytes } from 'ooxml-excel-editor/core'

const wb = jsonToWorkbook(
  [{ name: '张三', age: 25 }, { name: '李四', age: 30 }],
  { sheetName: 'People' },
)
writeFileSync('new.xlsx', await workbookToXlsxBytes(wb))
```

**Node headless 下不可用**(硬依赖浏览器 canvas / DOM,需在浏览器或 Electron 渲染进程跑):图片/PNG/JPEG 与 **PDF 导出**、`print()`、`downloadBlob`、内置 `DefaultEditor` 编辑器、`<ExcelViewer>` 组件渲染。另:`finalizeImages` 在 Node 会安全跳过(图片保留 `bytes`/`mime`,不生成 blob URL);URL 字符串入参走 `fetch`,**Node 用 `fs` 读 Buffer,别传本地路径字符串**。纯 Node 可用的导出面见 [EXTENDING.md](./EXTENDING.md#headless--node-安全-api-面)。

## 使用

### Vue

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { ExcelViewer } from 'ooxml-excel-editor'
import 'ooxml-excel-editor/style.css'

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

### React

同一套 core 引擎,React 薄壳。命令式 API 走 `ref`(`getSheetData` / `setSelection` / `downloadPdf` …,与 Vue 组件 ref 对齐):

```tsx
import { useRef, useState } from 'react'
import { ExcelViewer, type ExcelViewerHandle } from 'ooxml-excel-editor/react'

export function Preview() {
  const [file, setFile] = useState<File>()
  const viewer = useRef<ExcelViewerHandle>(null)
  return (
    <>
      <input type="file" accept=".xlsx" onChange={(e) => setFile(e.target.files?.[0])} />
      <ExcelViewer
        ref={viewer}
        src={file}
        fileName={file?.name}
        style={{ height: '100vh' }}
        onRendered={(wb) => console.log('已渲染', wb.sheets.length, '个工作表')}
        onSelectionChange={({ range, active }) => console.log(range, active)}
      />
    </>
  )
}
```

### Vue 2 (1.3.0+)

跟 Vue 3 1:1 复刻 (工具栏 / 公式栏 / dialog / 浮层 / events / API 全对齐). Vue 2 用 Options API + Composition API 都行:

```html
<template>
  <ExcelViewer
    ref="viewer"
    :src="src"
    :file-name="fileName"
    :editable="editMode"
    style="height: 100vh"
    @rendered="(wb) => console.log('已渲染', wb.sheets.length, '个工作表')"
    @cell-change="(p) => console.log(p.before.text, '→', p.after.text)"
  />
</template>

<script>
import ExcelViewer from 'ooxml-excel-editor/vue2'
import 'ooxml-excel-editor/style.css'
import 'ooxml-excel-editor/vue2.css'

export default {
  components: { ExcelViewer },
  data: () => ({ src: null, fileName: '', editMode: false }),
  methods: {
    download() { this.$refs.viewer.downloadXlsx() },
  },
}
</script>
```

完整 Vue 2 用法 + 跟 Vue 3 的差异表 见 [docs/Vue2.md](./docs/Vue2.md)。

### 仅引擎(不渲染 UI)

```ts
// 解析 + 读数据 + 导出,全程不依赖任何框架
import { parseWorkbook, loadArrayBuffer, getSheetData, WorkbookExporter } from 'ooxml-excel-editor/core'
```

### 全局注册(Vue 插件)

```ts
import OoxmlExcelPreview from 'ooxml-excel-editor'
import 'ooxml-excel-editor/style.css'

app.use(OoxmlExcelPreview) // 注册全局组件 <ExcelViewer />
```

### 读取数据(好用的数据访问 API)

不必自己再解析。两条路径,都拿同一份数据:

**A. 独立函数**(配 `parseWorkbook`,不渲染也能用):
```ts
import { parseWorkbook, loadArrayBuffer, getSheetData, sheetToJSON, getCellText, getWorkbookJSON } from 'ooxml-excel-editor'

const wb = await parseWorkbook(await loadArrayBuffer(file))
const sheet = wb.sheets[0]

getCellText(sheet, 1, 0, wb.date1904)         // 单格显示文本,如 '产品' / '¥1,234.50'
getSheetData(sheet, { date1904: wb.date1904 })            // 二维数组(显示文本)
getSheetData(sheet, { format: false, date1904: wb.date1904 }) // 二维数组(原始 number/Date)
sheetToJSON(sheet, { headerRow: 0, date1904: wb.date1904 })   // 首行作表头 → [{ 产品:'鼠标', 单价:89 }, ...]
getWorkbookJSON(wb)                            // 全簿 → { 表名: 对象数组 }(自动带 date1904)
```

**B. 组件 ref**(自动带 `date1904` + 默认当前表;插件 `ctx.viewer` 同样可用):
```ts
const viewer = ref()  // <ExcelViewer ref="viewer" />
viewer.value.getCellText(1, 0)                 // 显示文本
viewer.value.getCellValue(1, 0)                // 原始值
viewer.value.getSheetData()                    // 当前表 2D(默认显示文本)
viewer.value.getSheetJSON({ headerRow: 1 })    // 对象数组
viewer.value.getRangeData(viewer.value.getSelection())  // 取"我选中的"区域
```

| 函数 / 方法 | 返回 |
|---|---|
| `getCellValue` / `getCellText` / `getCellStyle` / `getCell` | 单格 原始值 / 显示文本 / 解析样式 / 模型 |
| `getSheetData` | 二维数组(稠密 `rows×cols`) |
| `getRangeData(range)` | 区域二维数组 |
| `sheetToJSON` | 对象数组(首行作 key,空表头回退列字母,全空行跳过) |
| `getWorkbookJSON` | `{ 表名: 对象数组 }` |

- **值 vs 文本**:`format` 默认 `true` → 套了数字/日期格式的**显示文本**(所见即所得);`{ format: false }` → 原始 `number/Date/boolean`。
- **合并单元格**:2D/JSON 里**锚点(左上)持值,其余为空**(单格 `getCellValue` 仍返回模型里的字面值)。
- **公式**:不重算,沿用 Excel 缓存结果(公式串见 `cell.formula`)。
- 也 re-export 了底层 `formatValue` / `cellKey`。

> 想要更底层的渲染模型,仍可直接用 `parseWorkbook` 的返回值 / `getWorkbook()` / `@rendered`(`WorkbookModel`:`sheets[].cells: Map<"row:col">`、`styles[styleId]`)。

## UI 区域速查(给调用方 & 二开者)

`<ExcelViewer>` 渲染出的 chrome 自上而下分这几块,每块都有**名字 / DOM 选择器 / 文件 / 替换方式**,改样式或替换某块时按这张表找位置。

```
┌─ <ExcelViewer> .excel-viewer (Vue) / .rxl (React) ─────────────────────────┐
│ ① Header (顶栏)            <ViewerToolbar>             ← #header slot 整条替换│
│    文件名 · 表数 · 缩放 · 导出 ▾                                             │
├────────────────────────────────────────────────────────────────────────────┤
│ ② ActionToolbar (操作栏)   .action-toolbar             ← :toolbar 配置/插件   │
│    [查找][筛选][复制][自动换行][冻结][缩放][导出] … 自动「⋯ 更多」           │
├────────────────────────────────────────────────────────────────────────────┤
│ ③ FormulaBar (公式栏 Fx)   .formula-bar                ← (无 slot,可全局 CSS)│
├────────────────────────────────────────────────────────────────────────────┤
│ ④ RenderArea (网格区)      .render-area + canvas.grid-canvas              │
│    ├ OverlayManager        图片/图表/形状/插件 overlay(DOM 叠加,随滚动)    │
│    ├ ContextMenu           .ooxml-context-menu(body 级,框架无关)          │
│    ├ LightboxHost          .ooxml-lightbox(body 级,图片放大+下载)         │
│    └ CellEditorHost        .editor-slot(自定义/内置 cell 编辑器)           │
├────────────────────────────────────────────────────────────────────────────┤
│ ⑤ SheetTabs (表标签)       .sheet-tabs                                    │
└────────────────────────────────────────────────────────────────────────────┘
```

| 区域 | DOM / class | 文件(壳) | 替换 / 自定义方式 |
|---|---|---|---|
| ① Header 顶栏 | `<ViewerToolbar>` (Vue) | [components/ViewerToolbar.vue](src/components/ViewerToolbar.vue) | `<template #header>` slot 整条替换;插槽签名见 README「分层 UI」 |
| ② 操作栏 | `.action-toolbar` | [components/ActionToolbar.vue](src/components/ActionToolbar.vue) / [react/RxlActionToolbar.tsx — 暂无,React 用内置按钮行]<br>builtin 渲染在 [components/ExcelViewer.vue](src/components/ExcelViewer.vue) `builtinTool()` | `:toolbar="[...]"` 数组 prop(内置 id + 自定义 `ToolbarItem`);插件 `definePlugin({ toolbar })` 追加项 |
| ③ 公式栏 | `.formula-bar` / `.rxl-formula-bar` | [components/ExcelViewer.vue](src/components/ExcelViewer.vue) / [react/ExcelViewer.tsx](src/react/ExcelViewer.tsx) | 暂无 slot;CSS 直接覆写(`.excel-viewer .formula-bar`)。命令式:`getCellEditString` / `commitActiveCellValue` |
| ④ 网格区 canvas | `.render-area > canvas.grid-canvas` | core: [render/canvas-renderer.ts](src/core/render/canvas-renderer.ts) | 渲染钩子 `cellStyle` / `transformModel` / 插件 `overlay` 返回 DOM;`:theme` 改配色 |
| ④a Overlay 层 | `.overlay-host` 四象限(主/冻结行/冻结列/冻结角) | core: [viewer/overlay-manager.ts](src/core/viewer/overlay-manager.ts) | `<template #overlay>` slot (Vue) / `overlay` prop 返回 DOM (React)/ 插件 `overlay` |
| ④b 右键菜单 | `.ooxml-context-menu`(body 级) | core: [edit/context-menu.ts](src/core/edit/context-menu.ts) | **Plan C 三层开放**:`:contextMenu` prop(`false` 关闭 / `(ctx,items)=>MenuItem[]` transform) · `@before-context-menu`/`@context-menu` 事件(`preventDefault()` 接管渲染) · 命令式 `openContextMenu(x,y,items?)` / `closeContextMenu()`;插件 `definePlugin({ contextMenu })` 链式贡献 |
| ④c 图片灯箱 | `.ooxml-lightbox`(body 级) | core: [viewer/lightbox-host.ts](src/core/viewer/lightbox-host.ts) | `imageLightbox={false}` 关闭;命令式 `openImageLightbox(src,fileName?)` |
| ④d 单元格编辑器 | `.editor-slot` | core: [edit/editor-host.ts](src/core/edit/editor-host.ts) + 内置默认 [edit/default-editor.ts](src/core/edit/default-editor.ts) | `:editor` prop = `EditorResolver`(按格返回工厂);插件 `editor`;详见 README「自定义编辑器」 |
| ⑤ 表标签 | `.sheet-tabs` | [components/SheetTabs.vue](src/components/SheetTabs.vue) / 内置于 React 壳 | 暂无 slot;`.excel-viewer .sheet-tabs` CSS 覆写 |

**重要:demo 顶栏不是组件的**。仓库里 `npm run dev` 看到的绿色顶栏(`.app-bar`)是 [src/App.vue](src/App.vue) 的 demo 框架栏(选 xlsx / 加载示例 / 编辑模式开关 / 演示按钮),React demo 同形见 [src/react-demo/main.tsx](src/react-demo/main.tsx)。它独立于 `<ExcelViewer>`,**用户 import 组件时不会出现**。

## API

### `<ExcelViewer>`

| Prop | 类型 | 说明 |
|---|---|---|
| `src` | `File \| Blob \| ArrayBuffer \| Uint8Array \| string(URL)` | 要预览的 .xlsx 数据源 |
| `workbook` | `WorkbookModel \| JsonInput` | **JSON 直渲(P3)** 优先于 `src`。WorkbookModel 直用;JsonInput 走 `jsonToWorkbook` 自动构造:① `unknown[][]` 二维数组 ② `Record<string,unknown>[]` 对象数组(首行表头) ③ `{ sheets:[...] }` 多表 |
| `jsonOptions` | `JsonLoadOptions` | JSON 直渲选项(`workbook = JsonInput` 时生效):`headerRow` / `sheetName` / `autoInfer`(数字串→数字、ISO 日期串→Date,默认 on) |
| `templateFile` | `ExcelSource` | **模板样式 overlay(P3 重设计 2026-06-08)** 一份 .xlsx 当**样式捐赠者** —— 模板贡献 styling(styles / merges / 列宽 / 行高 / freeze / theme),JSON / CSV 数据**在 A1 自然位置渲染**,模板的 raw 文字 / 占位符 / 图 / 图表 / 条件格式 **全部丢弃**。**仅在 `:workbook` (JSON / 模型) 数据源下生效**;`:src` (xlsx) 数据源自带格式,给 `:templateFile` 会被忽略并 console.warn。工具栏内置 `template` 项可在运行时切换 / 导入 / 清除 |
| `templateName` | `string` | 模板显示名(标题栏 `· 模板: xxx` 后缀);不给则取运行时 File.name。`:fileName` 同步:JSON 源未给名时默认 "JSON 数据" |
| `exportProgress` | `boolean`(默认 `true`)| **内置导出进度遮罩(P1.5)** 调 `viewer.downloadPdf` / `downloadImage` / `downloadXlsx` / `print` / 选区图片批量转换 时,壳自动建 `AbortController` + 接 `onProgress` → 显示居中模态(stage 标签 + 进度条 + 取消)。**关闭**用 `false`(纯回调走 `opts.onProgress`/`signal`);**自渲染**用 `#export-progress` 插槽(Vue)/ `renderExportProgress` (React) |
| `contextMenu` | `false \| ContextMenuTransform` | **右键菜单(Plan C)** — 三层开放:① 默认 = 内置菜单(editable 时弹) ② `false` = 不弹内置(`@before-context-menu` / `@context-menu` 事件仍触发,自渲染) ③ `(ctx, items) => MenuItem[] \| undefined` = transform 回调,在内置 items 上加 / 减 / 重排 |
| `fileName` | `string` | 标题栏显示的文件名(可选) |
| `editable` | `boolean` | 开启编辑(默认 `false` = 只读,行为与历史一致) |
| `pivotTable` | `boolean` | 透视表功能开关(默认 `false` = 关闭)。开启后(还需 `editable`):工具栏 `pivot-table` 入口可见、`createPivotTable`/`openPivotTableDialog` 等 API 生效、导出 .xlsx 回注真实 OOXML 透视表零件(overlay 模式同时保留原文件透视表) |
| `conditionalFormat` | `boolean` | 条件格式编辑开关(默认 `false` = 关闭、只读渲染)。开启后(还需 `editable`):工具栏 `conditional-format` 入口可见、`openConditionalFormatDialog`/`addConditionalRule`/`updateConditionalRule`/`removeConditionalRule`/`setConditionalRules`/`getConditionalRules` API 生效、导出 .xlsx 回写条件格式(overlay 保留原件未编辑规则原样,只增改用户改的)。支持全 6 类规则(cellIs / 公式 / 色阶 / 数据条 / 图标集 / top10)新建·编辑·删除,整体单次撤销 |
| `cellReadOnly` | `(cell, pos) => boolean` | 按格只读判定(编辑时) |
| `readOnlyRanges` | `MergeRange[]` | 只读区域(命中即只读,黑名单) |
| `editableTargets` | `EditableTarget \| EditableTarget[]` | **可编辑白名单**(2026-06-08)— 设了就是白名单语义:默认只读,**只**命中**任一** target 的格可编辑。4 种 target 形状自动识别:`{row,col}` 单格 / `{row}` 整行 / `{col}` 整列 / `MergeRange` 矩形。单值或数组都行,允许**不相邻**多 target。`undefined` (不传) = 不启用白名单 = 老行为(默认全可编辑);`[]` (显式空) = 全只读。与 `readOnlyRanges` / `cellReadOnly` 叠加 — 白名单命中后仍可被二次"黑"掉。运行时改:命令式 `viewer.setEditableTargets(targets)` |
| `strictDimensions` | `boolean` | **严格尺寸闸门**(Phase B, 2026-06-08)— 默认 `false`:`setColumnWidth` / `setRowHeight` / `autoFit` 仅受全局 `editable` 控制。设 `true` + 启用了 `editableTargets` → 该列/行至少有 1 格在白名单内才能改尺寸,否则 skip + emit `permission-denied` (`reason='dimension'`)。 |
| `readOnlyCellStyle` | `boolean \| CellStyleOverride \| CellStyleFn` | **只读视觉钩子**(Phase C, 2026-06-08)— 默认 `false` 无视觉差异(老行为);`true` 套内置浅灰底 `#f5f7fa`;对象 = 固定样式给所有只读格;函数 = 按格自定义。仅在该格 `editable=false` 时套用,跟 `editableTargets` 配合一眼看出哪些格可编辑。**鼠标光标**: 编辑模式下悬停只读格自动变 `not-allowed`(内置,不可关)。 |
| `editor` | `EditorResolver` | 自定义单元格编辑器工厂(返回任意 DOM) |
| `recalc` | `boolean` | 公式重算(默认 `false`;需 `editable`) |
| `formulaEngine` | `FormulaEngineFactory` | 自定义/自研公式引擎(默认 = 内置 MIT 引擎;可注入 `hyperFormulaEngineFactory` 或自研) |
| `pasteBehavior` | `Partial<PasteBehavior>` | 粘贴行为(默认 = 覆盖式 1:1)。逐项可配:`cellStyle`/`fill`(`overwrite`/`merge`/`skip`)·`rowHeight`(`source`/`keep`)·`colWidth`(`firstRowOnly`/`source`/`keep`)·`sourceMerges`(`apply`/`skip`)·`targetMerges`(`clear`/`keep`,默认清=修旧合并吞列)·`images`(`apply`/`skip`)。缺项回落默认。也可运行时 `viewer.setPasteBehavior(cfg)` / 右键「选择性粘贴」/ 工具栏「⚙ 粘贴配置」面板 |
| `readOnlyPrompt` | `'dialog' \| 'toast' \| 'none'` | 粘贴撞只读格的内置提醒(默认 `'dialog'`):`dialog` 弹窗**列出具体哪些格**只读 / `toast` 顶部气泡 / `none` 只发 `permission-denied` 事件。逐格精确(编辑模式下也可能有只读格) |
| `cellImageFit` | `'fill' \| 'contain' \| 'cover'` | WPS 单元格内嵌图贴合方式(默认 `contain` 等比,与 WPS 渲染一致) |
| `imageLightbox` | `boolean` | 图片点击放大灯箱(默认 `true`;只读单击图放大、编辑右键「查看大图」) |
| `toolbar` | `false \| Array<string \| ToolbarItem>` | 操作工具栏配置(默认 `['find','filter','sort']`)。内置 id 见 [操作工具栏](EXTENDING.md#操作工具栏可配置--可插件--响应式);`false` 不渲染。可混入自定义项 |
| `plugins` | `ExcelPlugin[]` | 插件数组(`definePlugin` 打包 theme/cellStyle/transformModel/events/overlay/toolbar/setup);见 [插件](EXTENDING.md#插件-defineplugin) |
| `openLinks` | `boolean` | 单击超链接是否自动打开(默认 `true`;`false` 只发 `@hyperlink` 事件,自己处理) |

> 编辑相关 props 详见下方 [编辑](#编辑可选默认只读) 章节;数据验证(开 `editable` 后**编辑时拦截非法输入** + 列表型下拉选值)无需额外 prop,解析到的规则自动生效。

| 事件 | 载荷 | 触发时机 |
|---|---|---|
| `rendered` | `WorkbookModel` | 解析并首次渲染完成 |
| `error` | `string` | 解析失败(友好文案) |
| `cell-change` | `{ before, after, source }` | 单元格变更(编辑/撤销/重做/公式级联);`before`/`after` 是**完整快照**(底层 cell + 解析 style + raw/computed/text) |
| `edit-start` / `edit-commit` | `{ cell, ... }` | 进入 / 提交编辑 |
| `dim-change` | `{ axis, index, before, after }` | 列宽/行高变更 |
| `image-change` | `{ index, before, after }` | 图片增删移改(前后 `ImageAnchor`) |
| `struct-change` | `{ op, at, count }` | 增删行列 |
| `dirty-change` | `{ dirty }` | 有/无未保存修改 |

容器需要有明确高度(组件填满父容器),例如 `style="height: 100vh"`。

### 具名导出

> **四个入口同源**:`ooxml-excel-editor`(Vue 3)、`/react`、`/vue2`、`/core` 都 **re-export 同一套框架无关 core 公共 API** —— 任一入口都能拿到下表全部(各入口再各自加自己的组件)。所以 React/Vue2 用方不必绕到 `/core`。

| 导出 | 说明 |
|---|---|
| `ExcelViewer` | 预览/编辑组件(各框架入口各自的) |
| `default` | Vue 3 / Vue 2 入口默认导出 = Vue 插件(`app.use`) |
| **解析/加载** | `openWorkbook(src)`(一行门面 = 归一化 + 解析,Node/浏览器通用)· `parseWorkbook(buffer)`(`ArrayBuffer\|Uint8Array → Promise<WorkbookModel>`,优先 Worker)· `loadArrayBuffer(src)`(多种输入归一)· `jsonToWorkbook(data)`(数据直建模型) |
| **读数据 API** | `getCellValue` / `getCellText` / `getCellStyle` / `getSheetData` / `getRangeData` / `sheetToJSON` / `getWorkbookJSON` / `cellDisplayText` |
| **格式/工具** | `formatValue`(数字格式)· `cellKey` · `colIndexToLetters` |
| **插件 / 主题** | `definePlugin` · `DEFAULT_THEME` / `mergeTheme` |
| **公式引擎**(1.14.0) | `builtinFormulaEngineFactory`(默认,MIT)· `hyperFormulaEngineFactory`(HyperFormula,GPL/商业)· `FUNCTION_NAMES`(已支持函数名)· `BuiltinFormulaEngine` |
| **导出工具** | `workbookToXlsxBytes`(→ `Uint8Array`,**纯 Node 落盘用**)· `workbookToXlsxBlob`(→ `Blob`,浏览器下载用)· `toCsv` / `toWorkbookJson` · `canvasToBlob` / `canvasToDataURL` / `downloadBlob`(后三者需浏览器) |
| **类型** | `WorkbookModel` / `SheetModel` / `CellModel` / `CellStyle` / `CellStyleOverride` / `MergeRange` / `ConditionalRule` / `DataValidationRule` / `ChartSpec` / `ImageAnchor` / `PivotTableModel` / `CssColor` / `ExcelSource` / `ViewerApi` / `ExcelPlugin` / `FormulaEngine` / `FormulaEngineFactory` / `EditConfig` / 导出选项类型(`PdfExportOptions`/`ImageExportOptions`/…)等 |

> 想看完整出口清单见 [`src/core/index.ts`](src/core/index.ts);深度二开(直接用 `ViewerController` / `CanvasRenderer` / `EditController` / 模型 mutations 等内部件)也都从这些入口导出,见 [ARCHITECTURE.md](ARCHITECTURE.md)。

## 编辑(可选,默认只读)

组件默认是**只读预览**,行为与历史完全一致、零额外成本。传 `:editable` 才进入编辑模式,所有编辑能力建在**一份可变内存模型**(`WorkbookModel`)上,读 / 写 / 事件 / 导出共用同一层。

### 开启 + 只读控制

```vue
<ExcelViewer
  :src="src"
  :editable="true"
  :read-only-ranges="[{ top: 0, left: 0, bottom: 0, right: 4 }]"   <!-- 表头整行只读 -->
  :cell-read-only="(cell, pos) => pos.col === 0"                    <!-- 第 0 列只读 -->
  @cell-change="onCellChange"
/>
```

开编辑后:**双击 / F2 / 直接打字**进格编辑(Enter 提交下移、Tab 右移、Esc 取消、Shift+Enter 插换行),**Ctrl+Z / Ctrl+Y** 撤销重做,拖列头/行头边界改宽高,拖浮动图片移动。只读格 / 区域不进编辑。

**长文本编辑(WPS 风格, 1.2.1)**:默认编辑器是 `<textarea>`, 输入长文本自动**换行 + 向下浮起撑高**, 编辑期显示完整内容;提交后单元格行高保持原样(跟 WPS 一致, 不持久化撑高;如需永久保持多行, 单元格设 `wrapText=true`)。自定义编辑器可通过 `CellEditorReturn.getDesiredHeight(width)` 复用此撑高机制 + `ctx.reposition()` 主动触发重撑。

### 命令式编辑 API(模板 ref / 插件 `viewer`)

| 类别 | 方法 |
|---|---|
| 值 | `editCell(row,col,value)` · `editRange(range,values[][])` · `clearRange(range)` |
| 粘贴 | `pasteText(tsv,at?)` · `pasteRichHtml(html,at?,behaviorOverride?)`(Excel/WPS 复制 → 字体/颜色/填充/边框/对齐/换行/合并/内嵌图 + 单次撤销) · `pasteImageBlob(blob,at?)`(单图/拖文件落格) |
| 粘贴行为 | `getPasteBehavior()` · `setPasteBehavior(cfg)`(改 Ctrl+V/右键默认)· `openPasteConfigDialog()`(弹「⚙ 粘贴配置」面板,框架无关 DOM 三壳共用)。覆盖式 1:1(默认)/ 合并 / 仅值;`targetMerges:'clear'` 默认清目标旧合并防吞列;`colWidth:'firstRowOnly'` 仅首行取源宽。右键「选择性粘贴 ▸ 覆盖格式 / 保留原样式(仅值)」逐次选 |
| 样式 | `setStyle(range, patch)`(`patch` = `CellStyleOverride`:font/fill/borders/对齐/numFmt) |
| 背景/字体色 | `getActiveFillColor()` · `getActiveFontColor()`(回显活动格当前色 #RRGGBB) · `setSelectionFill(color\|null)`(null=清除填充) · `setSelectionFontColor(color)` |
| 自动换行 | `getSelectionWrapState()`→`'all'\|'none'\|'mixed'`(工具栏 active 用) · `toggleWrapTextOnSelection()`(WPS 风格 toggle:全开/全关;行高按内容重撑,只扩不缩) |
| 列宽行高 | `setColumnWidth(col,px)` · `setRowHeight(row,px)` |
| 行列结构 | `insertRows(at,count?)` · `deleteRows(at,count?)` · `insertCols(at,count?)` · `deleteCols(at,count?)` |
| 图片 | `getImages()` · `addImage(anchor)` · `removeImage(i)` · `moveImage(i,dxPx,dyPx)` · `resizeImage(i,wPx,hPx)` |
| WPS 内嵌图 | `getCellImages()` · `setCellImageFit('fill'\|'contain'\|'cover')` · `convertImageToCellAuto(imgIdx)`(就近) · `convertAllImagesToCells(col?)`(整表/整列) · **`convertImagesInRangeToCell(range)`**(选区批量,单次撤销) · `convertCellImageToFloat(row,col,size?)`(嵌入→浮动) · **`convertCellImagesInRangeToFloat(range, size?)`**(选区批量浮动化) · `convertImageToCell(imgIdx,row,col)`(显式目标) |
| 图片放大 | `openImageLightbox(src,fileName?)`(命令式弹大图+下载) · `getCellImageAt(row,col)`(某格是否内嵌图→`{id,src,mime}`) |
| 撤销/进编辑 | `undo()` · `redo()` · `canUndo()` · `canRedo()` · `beginEdit(row,col)` · `cancelEdit()` · `isEditing()` · `getEditingCell()` |
| 公式栏 | `getCellEditString()`(活动格可编辑字符串:公式→`=…`,数值→原始数字串) · `canEditActiveCell()` · `commitActiveCellValue(value, move?)`(顶部 Fx 公式栏可编辑并与单元格联动,底层即用这套) |
| 查询/状态 | `getCellSnapshot(row,col)` · **`inspectCell(row,col)`**(全息体检:snapshot + 合并区 + 浮动图覆盖 + WPS 内嵌图 + 数据验证 + 条件格式命中 + 链接/批注) · `isDirty()` · `resetToOriginal()` · `isRecalcReady()` · `getVirtualExtent()`(当前虚拟行列范围) |
| 导出 | `exportXlsx/downloadXlsx` · `exportJson/downloadJson` · `exportCsv/downloadCsv`(见 [导出](EXTENDING.md#导出--打印)) |

所有写操作(含拖拽改宽高/移图)**统一进撤销栈**、发对应事件、翻**脏标记**;`resetToOriginal()` 一键放弃全部修改、还原到刚加载的原件。

### 自定义编辑器 `:editor`

返回一个工厂(拿到 `ctx` 里的快照 / 矩形 / `commit`/`cancel`),挂任意 DOM 当编辑控件 —— 下拉、日期选择器、图片选择器、带按钮的面板…… 是框架无关的 DOM(Vue/React 共用)。`commit` 可只给值,也可 `{ value, style }` 同时套样式。

```ts
const myEditor: EditorResolver = (cell, pos) => {
  if (pos.col !== 0) return                       // 该列不接管 → 用内置文本编辑器
  return (ctx) => {                               // 工厂:返回 DOM
    const sel = document.createElement('select')
    sel.innerHTML = `<option>A</option><option>B</option>`
    sel.value = String(ctx.snapshot.raw ?? '')
    sel.onchange = () => ctx.commit(sel.value)
    return sel
  }
}
```

插件也可经 `editor` 字段贡献(多插件数组序,组件 prop 最后覆盖),与 `cellStyle`/`overlay` 同范式。

### WPS 单元格内嵌图(DISPIMG)

不少 .xlsx 由 **WPS** 导出,其"嵌在单元格里的图"用了 WPS 私有存法(`xl/cellimages.xml` + 单元格 `=DISPIMG("id",1)` 公式),标准解析读不出来 → 普通工具打开会**缺图**。本组件:

- **自动识别并展示**:解析 `cellimages.xml` 登记表,把图**画进单元格内**(随行高列宽 / 滚动 / 裁剪 / 冻结 / 缩放),非浮动叠加。**贴合方式可配置** `cellImageFit`:`contain`(默认,等比缩放,**与 WPS 渲染一致**——WPS 打开导出文件时 DISPIMG 固定按 contain 显示)/ `fill`(拉伸铺满随格变形)/ `cover`(等比裁剪铺满)。`getCellImages()` 读登记表;`getCellSnapshot(row,col).cell.dispImgId` 看某格是否内嵌图。
- **一键浮动 ⇄ 嵌入互转**(编辑模式):
  - `convertImageToCellAuto(imgIdx)` —— **就近嵌入**:图压在哪个单元格上就嵌进哪格(几何反推,无需手指定目标格)。
  - `convertAllImagesToCells(col?)` —— **整表/整列批量**就近嵌入(`col` 给定只嵌该列),一次进撤销栈(单次 Ctrl+Z 全撤),返回嵌入张数。
  - `convertCellImageToFloat(row,col,size?)` —— 内嵌图拎回浮动图。
  - 右键单格菜单:「将此处浮动图嵌入单元格 / 整列浮动图嵌入(N 张)/ 整表浮动图嵌入(N 张)/ 内嵌图转为浮动图」。
  - 全部入撤销栈、发 `cell-change`/`image-change`、翻脏标记。(`convertImageToCell(imgIdx,row,col)` 仍保留,用于显式指定目标格。)
- **导出往返**:`downloadXlsx()` / `exportXlsx()` 导出时,在 ExcelJS 写出后**于 zip 层回注** WPS 私有件(`cellimages.xml` + rels + media + `[Content_Types].xml`/`workbook.xml.rels` 补丁,从模型重建),原有的 + App 内新转的内嵌图导出后用 WPS 打开都正常显示。rebuild / overlay 两种保真模式均覆盖;无字节的 blob-only 图除外。

### 不连续多区域选择(1.13.0)

**Ctrl/⌘ + 点击** 行头 / 列头 / 单元格 → 加选不相邻区域(Shift 仍是连续区间);普通点击 / 键盘导航回单选。多选时复制把各区**逐行堆叠**成块(TSV + HTML,粘到 Excel/WPS/app 内都成堆叠块),状态栏统计跨所有区聚合。`getSelectionRanges()` / `hasMultiSelection()` 读多选状态。纯框架无关 core 交互,三壳一致。

### 格式刷(1.12.0)

工具栏 `format-painter` 入口:先选**源格**点按钮采样其完整样式(字体/填充/边框/对齐/换行/数字格式),再**点或拖**目标格/区域即把格式刷上(单次撤销);`Esc` 或再点按钮退出,待刷时光标变 `copy`。也可 `startFormatPainter()` / `isFormatPainterArmed()` / `cancelFormatPainter()` 直调。纯框架无关 core 交互,三壳一致。

### 数字格式 / 批注 / 查找替换(1.11.0)

- **查找替换**:`Ctrl+F` 打开查找栏,编辑模式下多出替换行 —— 替换输入 + 「替换」(替换当前并跳下一个)/「全部替换」(整体单次撤销),支持区分大小写 / 全字匹配,跳过只读格。
- **数字格式编辑器**:工具栏 `number-format` 入口打开对话框(框架无关 DOM,三壳共用)—— 分类(常规/数值/货币/百分比/日期/时间/文本/自定义)+ 选项(小数位数 / 千分位 / 负数红色 / 货币符号 / 日期时间预设)+ 实时预览 + 可直接编辑格式代码 → 确定即套到选区(单次撤销)。也可 `setSelectionNumberFormat(code)` / `openNumberFormatDialog()` 直调。
- **批注编辑**:右键单格「插入/编辑/删除批注」打开对话框(多行文本 + 确定/删除/取消);`getCellComment` / `setCellComment(row,col,text)`(空串删除)/ `openCommentEditor()` 直调,单次撤销;**导出 .xlsx 回写批注**(rebuild + overlay)。

### 自动填充柄(拖拽填充序列)

开 `:editable` 后,选区右下角出现**填充柄**小方块(像 Excel/WPS)。拖它向下/上/左/右,松手即把源选区的模式**接续填充**进新增格(整体单次撤销):

- **全数值**:1 个 → 复制;≥2 个 → 等差外推(`1,2` → `3,4,5`;`2,4` → `6,8,10`,支持递减)
- **全日期**:1 个 → 每格 +1 天;≥2 个 → 按相邻差外推
- **"前缀+末尾整数"文本**:`Item 1` → `Item 2,Item 3`;`第1周` → `第2周`;`A01` → `A02`(保留前导零位宽)
- **星期 / 月份名**(中英常见写法):`周五` → `周六,周日,周一`;`Nov` → `Dec,Jan`(循环接续)
- **其它**:循环复制源值

**按住 Ctrl/⌘ 拖**翻转"复制 ↔ 序列"(对齐 Excel):单个数字普通拖=复制、Ctrl 拖=递增(`5`→`6,7,8`);两个数字序列普通拖=等差、Ctrl 拖=复制;日期/星期月份/文本递增普通=序列、Ctrl=复制。

主轴按鼠标偏移更大的方向决定(上下 vs 左右);列 / 行各自独立成序列;跳过只读格。注:v1 填充**值**,不复制源单元格的格式。纯框架无关 core canvas 交互,三壳一致。

### 公式重算(可换引擎)

开 `:recalc` 后,编辑公式格或被公式引用的格 → 依赖格**自动级联重算**,每个变动都发 `cell-change`(`source: 'api'|'ui'|'undo'|'redo'`)。

**默认引擎(1.14.0 起)= 内置 MIT 引擎**:从零实现的解析 + 求值 + 依赖图 + 拓扑级联 + 循环检测,**零依赖、无 GPL**,覆盖日常 ~60 个常用函数(SUM/AVERAGE/IF/IFERROR/VLOOKUP/INDEX/MATCH/SUMIF/COUNTIF/ROUND/LEFT/MID/CONCAT/DATE/TODAY… 见 `FUNCTION_NAMES` 导出)。函数集比 HyperFormula 小,但日常足够。

需要更全覆盖:`:formula-engine="hyperFormulaEngineFactory"` 注入 **HyperFormula**(可选 peer `npm i hyperformula`,GPL-3.0/商业双授权,~395 函数),或注入自研引擎(实现 `FormulaEngine` 接口)。`isRecalcReady()` 查引擎是否就绪。**打开文件的显示值与引擎无关**(读原件缓存结果),引擎只在编辑后重算时介入。

**公式自动补全**(1.14.0):在单元格里输 `=SU` 时下方弹**函数名列表 + 参数提示**;↑↓ 选、Enter/Tab 接受(插入 `SUM(` 并把光标移进括号)、Esc 关、点选即填。列表 = 引擎实际支持的函数(所见即所得),框架无关默认编辑器内置,三壳一致。

### 事件 = 前后完整快照

每次编辑都以 **`cell-change`** 通知:`{ before, after, source }`,`before`/`after` 是 `CellSnapshot` —— 不只 raw,还含**计算值 computed、显示文本 text、整个底层 `CellModel` + 解析后 `style`**。事件流**和**查询 API(`getCellSnapshot`)同一份底层结构 → JSON/CSV/XLSX 导出都复用它,无需按格式各写一遍解析。另有 `dim-change`/`image-change`/`struct-change`/`dirty-change`(见上方事件表)。

### v1 已知限制

- 增删行列**会自动重写公式引用**(`=A5` 上方插一行 → `=A6`,删被引用行 → `#REF!`,含跨表 `Sheet1!A5`)。
- 写回 .xlsx 默认从模型重建,丢 VBA/工作表保护/复杂 DrawingML 等;需要更高保真可用 `exportXlsx({ fidelity: 'overlay' })` **重载原件叠加编辑**(见 [导出保真边界](EXTENDING.md#导出--打印))。

## 扩展 / 二开

不改源码就能定制外观（`:theme`）、数据/渲染钩子（`transformModel`/`cellStyle`）、自定义编辑器（`:editor`）、右键菜单 transform、操作工具栏自定义项、分层 UI slots、命令式 API、导出/打印高级选项、以及**插件**（`definePlugin` 打包多种扩展点、跨框架可用）—— 完整 API 见 **[EXTENDING.md（二开 / 扩展 API 手册）](EXTENDING.md)**。

想了解内部结构 / 在哪改代码：见 **[ARCHITECTURE.md](ARCHITECTURE.md)**。

## 浏览器支持

现代浏览器(Chrome/Edge 80+、Safari 15+、Firefox 114+,需支持 Canvas / ResizeObserver)。

> **解析线程**:发布的组件库在**主线程**解析(`exceljs` 1.3.2+ 已内联进 dist chunks,运行时动态加载,不重复进你的产物)。本仓库的 demo/dev 额外启用了 **Web Worker** 解析(大文件不卡 UI)。如果你的应用要处理很大的文件,可直接用导出的 `parseWorkbook` 包进你自己的 Worker。

## 范围边界

- **编辑**已支持(默认只读;开 `editable`,见 [编辑](#编辑可选默认只读))。下列为暂不覆盖项:
- 增删行列**不自动重写公式引用**;写回 .xlsx 丢 VBA/工作表保护/复杂 DrawingML
- 透视表:已有透视表**数据按普通单元格显示**并显示只读字段按钮;`pivot-table` 入口可从当前选区选择生成位置,生成静态透视汇总表到当前表指定单元格或新建工作表,随后打开 WPS 风格右侧“数据透视表”字段面板。创建出的透视表**空白起步**(对齐 WPS/Excel:不自动猜字段),在右侧面板里选字段填充。面板支持搜索字段;**勾选字段列表的复选框** = 加入透视表(数值字段→值,其它→行),取消 = 移出;也可用 筛/列/行/Σ 按钮或拖拽把字段加入“筛选器 / 列 / 行 / 值”四区,并可拖到移除区删除字段。改动后即时重建结果:筛选器支持“全部 / 非空 / 多选(勾选要保留的具体值,WPS 风格)”,列区生成横向分组,行区生成纵向分组,值区可放多个字段且可切换“求和 / 计数 / 平均值 / 最大值 / 最小值”。**透视结果是“活”的**:① 编辑源数据区任意单元格(含撤销/重做)后,所有透视表自动按源区域重算;② 放两个及以上行字段时,外层分组带小计且可点行首 [−]/[+] 折叠/展开内层明细(单行字段为扁平结果,无折叠)。也可调用 `createPivotTable({ sourceRange, sourceSheetIndex, output, layout, showPanel })` 不经过页面直接创建;旧的 `createPivotTableFromSelection({ rowFieldIndex, valueFieldIndex, output })` 仍保留。创建出的 `PivotTableModel` 会保存 `source` 和 `layout` 元数据,供运行时重建与导出使用。失败时会提示原因。**整个透视表功能由 `pivotTable` 配置开启(默认关闭)**,关闭时入口/API/导出回注全部不生效、行为与历史版本一致。**导出 .xlsx 时回注真实 OOXML 透视表零件**(pivotCacheDefinition / pivotCacheRecords / pivotTableDefinition + 全套 rels,带 `refreshOnLoad`):Excel/WPS 打开导出件即识别为真透视表并按源区域重算原生布局;静态汇总结果仍在单元格里,不支持透视的查看器也能看。筛选器导出语义对齐 WPS:“= 具体值”写入页字段选中项(打开还原筛选状态),“多选/非空”映射为 `multipleItemSelectionAllowed` + 未选项隐藏(`item@h`),“全部”不写选中。原文件已有的透视表在 **overlay 导出模式**下从原件 zip 原样搬运整套零件(保持“打开→编辑→另存,透视表仍在”);rebuild 模式因结构可能被增删行列改动不搬运(退化为普通单元格)
- 条件格式编辑:解析时全 6 类规则(突出显示单元格 `cellIs` / 公式 `expression` / 色阶 `colorScale` / 数据条 `dataBar` / 图标集 `iconSet` / 项目选取 `top10`)都渲染;**开 `conditionalFormat`(还需 `editable`)后可新建/编辑/删除**。工具栏 `conditional-format` 入口打开**条件格式管理对话框**(框架无关 DOM,三壳共用一份,UI 天然 1:1):列出当前表所有规则(每条可「编辑」「删除」)+「新建规则」。新建时选规则类型,弹出对应编辑器 —— cellIs:运算符(大于/小于/介于/不介于/等于/不等于/大于等于/小于等于)+ 值(介于类两个值)+ 命中格式(填充色 / 字体色 / 加粗);expression:自定义公式 + 命中格式;colorScale:双色/三色 + 各色标取色;dataBar:条颜色 + 渐变开关;iconSet:7 种图标集(三色交通灯 / 三向箭头 / 三符号 / 三色旗 / 四等评级 / 五等评级 / 五象限)+ 反向;top10:前/后 + 个数 + 百分比开关 + 命中格式。**新建规则默认套到当前选区**;编辑保留原区域。所有改动**整体单次撤销**(`set-conditional` 命令替换整张规则集),改完即时重渲。也可不开对话框直接调 API:`getConditionalRules()` / `addConditionalRule(rule)`(返回新 id)/ `updateConditionalRule(id, patch)` / `removeConditionalRule(id)` / `setConditionalRules(rules)` / `openConditionalFormatDialog()`。**导出 .xlsx 回写条件格式**:`rebuild` 与 `overlay` 都回写(1.9.0 起 rebuild 不再丢条件格式);**从文件解析、未在 app 内编辑过的规则按原始 OOXML 原样回写**(保留色阶/数据条/图标集的 cfvo 阈值等不全建模字段,零退化),用户新建或编辑过的规则按模型重建(全 6 类)。整个功能由 `conditionalFormat` 配置开启(默认关闭),关闭时入口/API/导出回写全部不生效、行为与历史一致。当前限制:色阶/数据条/图标集的**阈值(cfvo)** 暂不在编辑器里微调(未编辑时靠原样回写保住,用户新建用默认阈值);文本包含等 OOXML 子类型用 `cellIs`/`expression` 表达。
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

> **e2e 说明**:`npm run test:e2e` 用 Playwright 起 dev 服务 + 无头 Chromium,加载示例 → 渲染 → 导出 PNG/位图PDF/矢量PDF,校验产物(PNG 魔数、`%PDF`、矢量 PDF 的文字操作符数量多于位图)。覆盖 node 单测做不到的真实 canvas/jsPDF 绘制。首次需 `npx playwright install chromium` 下载浏览器(本仓库 `@playwright/test` 固定 `1.58.0` 对应 chromium-1208)。**Vue 3 / React demo 在 5300**(`/` 与 `/react.html`),**Vue 2 demo 在独立 5302**(plugin-vue2 SFC 隔离);两个 dev server 都由 Playwright 自动拉起。三壳均有 e2e 覆盖。

## 文档 / 二开

- **[EXTENDING.md](./EXTENDING.md) —— 二开 / 扩展 API 手册**(主题 `:theme` / 数据·渲染钩子 / 自定义编辑器 / 右键菜单 transform / 工具栏自定义项 / 分层 UI slots / 命令式 API / 导出·打印高级选项 / 插件 `definePlugin`)
- [ARCHITECTURE.md](./ARCHITECTURE.md) —— 包/入口、core 分层、数据流、`ViewerController` 桥接、"加功能改哪"
- [CONTRIBUTING.md](./CONTRIBUTING.md) —— 本地跑通、改动流程、不可破坏的硬约束
- [CHANGELOG.md](./CHANGELOG.md) / [RELEASING.md](./RELEASING.md) —— 变更记录 / 发布清单
- [docs/编辑权限与只读边界.md](./docs/编辑权限与只读边界.md) —— **EditableTarget 白名单 / DimTarget 尺寸多形态 / readOnlyCellStyle 视觉钩子 / permission-denied 事件** 体系化说明(1.2.0)
- [docs/Vue2.md](./docs/Vue2.md) —— **Vue 2 兼容子入口**完整文档(1.3.0;`ooxml-excel-editor/vue2` 跟 Vue 3 / React 壳 ~100% 功能对齐)

> **React props/events** 与 Vue 对齐(事件用 camelCase 回调:`onRendered`/`onError`/`onCellClick`/`onSelectionChange`/`onSheetChange`/`onHyperlinkClick`),命令式句柄 `ExcelViewerHandle` 与 Vue 组件 ref 同名方法一致。[EXTENDING.md](./EXTENDING.md) 里的**插件 `definePlugin`** 目前服务 Vue 壳;React 壳已可用全部 props/命令式 API/事件,插件 overlay 跨框架化在路线图中。

## License

MIT

---

<a name="english"></a>
## English

**🔗 Live demo: https://ojadeveloper.github.io/ooxml-excel-editor/**

A **Vue 3 + Vue 2 + React high-fidelity `.xlsx` preview & edit component** with a from-scratch parser and canvas renderer. Renders cells, number formats, merges, conditional formatting, images, charts (via ECharts), sparklines, comments, data validation, frozen panes, and supports selection / copy / hyperlinks. **Read-only by default**; set `editable` to enable editing — cell values / styles / column-row sizes / floating images / insert-delete rows-cols, with undo-redo, before/after full-snapshot events, dirty tracking + reset-to-original, swappable formula recalc engine, custom cell editors, and **export back to .xlsx / JSON / CSV**. Parsing runs in a Web Worker (with main-thread fallback). Only the framework (`vue` / `react`) is a peer dependency; `exceljs` / `fflate` / `jspdf` / `hyperformula` are bundled into `dist` (no manual install).

```bash
npm i ooxml-excel-editor vue
```

```ts
import { ExcelViewer } from 'ooxml-excel-editor'
import 'ooxml-excel-editor/style.css'
```

**Headless / Node (no browser):** use `ooxml-excel-editor/core` to parse `.xlsx`, extract data, edit and write back — no canvas needed. `npm i ooxml-excel-editor` only.

```ts
import { readFileSync, writeFileSync } from 'node:fs'
import { openWorkbook, sheetToJSON, setCellValue, workbookToXlsxBytes } from 'ooxml-excel-editor/core'

const src = readFileSync('input.xlsx')
const wb = await openWorkbook(src)              // accepts a Node Buffer directly
const rows = sheetToJSON(wb.sheets[0])          // array of row objects (display text)
setCellValue(wb.sheets[0], 1, 2, 123.45)
writeFileSync('out.xlsx', await workbookToXlsxBytes(wb, {
  fidelity: 'overlay',
  sourceBuffer: src.buffer.slice(src.byteOffset, src.byteOffset + src.byteLength),
}))            // workbookToXlsxBytes → Uint8Array (not a Blob)
```

Not available headless (need a browser canvas/DOM): image/PDF export, `print()`, the built-in cell editor, and component rendering. See the 中文 [Node 用法](#node--服务端-headless-用法) section for details.

See the API table above. MIT licensed.
