# ooxml-excel-editor

> Vue 3 + **Vue 2** + React 高保真 **.xlsx 预览 / 编辑组件** —— Canvas 渲染,**默认只读预览,可选开启编辑**。从零实现解析与渲染,尽量还原微软 Excel 打开工作簿的观感。**三个壳 UI 1:1 对齐**(Vue 3 SFC 是标准,Vue 2 / React 复刻)。

[English](#english) · 中文

## ⚡ 快速开始

**装**(按框架二选一):

```bash
npm i ooxml-excel-editor vue exceljs                  # Vue 3 (默认入口)
npm i ooxml-excel-editor react react-dom exceljs      # React 壳 (/react 子入口)
npm i ooxml-excel-editor vue@2.7 @vue/composition-api exceljs  # Vue 2.6/2.7+ (/vue2 子入口, 1.3.0+)
```

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
- 🌈 **条件格式**:色阶 / 数据条 / 图标集 / cellIs / top10
- 🖼 **图片 + 图表**(DrawingML → ECharts 近似还原)、**形状/文本框**、**迷你图**(sparklines)、**批注**、**数据验证**下拉、**自动筛选**样式
- 📌 **WPS 单元格内嵌图(DISPIMG)**:识别并展示 WPS 私有的"嵌在格里的图"(普通工具会缺图);编辑模式下支持**一键浮动 ⇄ 嵌入互转**。见 [WPS 单元格内嵌图](#wps-单元格内嵌图dispimg)
- 🔍 **图片点击放大 + 下载原图**:网格里的图(内嵌图/浮动图)点开看大图、下载原件。只读模式单击图放大、编辑模式右键「查看大图」。`imageLightbox` prop 控制(默认开),`openImageLightbox(src)` 命令式打开。
- 📋 **从 Excel/WPS 富粘贴**:`Ctrl+V` 解析剪贴板 HTML → 完美还原字体/颜色/填充/边框/对齐/合并单元格,整块单次撤销。图片走多通道(data-uri / 单图 / 拖文件);**注**:Excel 区域复制的内嵌图进不了浏览器剪贴板(浏览器限制)。
- 📝 **文本溢出**到相邻空格、**自动行高**
- 🖱 **交互**:单元格选区(合并感知)、拖选、公式栏、状态栏(计数/求和/均值/最值)、超链接可点、裁切文本悬停看全文、Ctrl+C 复制为 TSV、**Ctrl+F 查找**(高亮 + 上/下定位 + 计数 + 区分大小写/全字匹配)、**自动筛选**(点下拉真能筛:去重值多选 + 搜值 + 清除)
- 🖨 **导出 / 打印**:整表/选区/多表导出 **PNG/JPEG**、**PDF**(位图 + **矢量·文字可选可搜**两种)、**系统打印**(可另存 PDF);默认还原原生 `pageSetup`(纸张/方向/页边距/缩放/打印区域/**打印标题行列每页重复**);宽表**横向跨页**(页矩阵);`beforeRenderPage` 注入页眉/页脚/水印、`configureDoc` 注册字体;内置「导出设置」对话框
- ⚡ **按需加载**(无图表文件不下载 echarts、不导出 PDF 不下载 jspdf)、**友好错误兜底**(损坏/加密/旧 .xls)、解析失败自动给出可读提示

- 📤 **数据读取 API**:不必自己再解析 —— `getCellText`/`getSheetData`/`sheetToJSON`/`getRangeData`(独立函数 + 组件 ref 方法),值/显示文本可选,合并/日期/数字格式都处理好
- ✏️ **编辑(可选,默认只读)**:开 `editable` 即可编辑 —— 单元格值 / 样式(粗体/对齐/填充)/ 列宽行高 / 浮动图片(拖拽移改)/ 增删行列;**撤销重做**(Ctrl+Z/Y)、**前后完整快照事件**、**脏状态 + 一键还原原件**;可换**公式引擎**自动重算依赖格;可注入**自定义编辑器**(下拉/日期/图片选择器);**导出回 .xlsx / JSON / CSV**(所见即所得)。见 [编辑](#编辑可选默认只读)

> 纯预览不需要公式引擎 —— .xlsx 缓存了公式结果,直接显示;仅开启**编辑 + 重算**时才用(可选 `hyperformula`)。详见 [EXCEL还原难点.md](./EXCEL还原难点.md)。

## 安装

一个包,**四个子入口** —— 框架无关的 core 引擎被 Vue 3 / React 两个壳共享(`dist/core.js` 只打一份),Vue 2 因 SFC 编译器跟 Vue 3 冲突独立打包(内嵌 core)。按你的框架装对应 peer:

```bash
# Vue 3 项目
npm i ooxml-excel-editor vue exceljs

# React 项目
npm i ooxml-excel-editor react react-dom exceljs

# Vue 2.6.x 或 2.7+ 项目 (1.3.0+) — 必装 @vue/composition-api (兼容 2.6 + 2.7)
npm i ooxml-excel-editor vue@2.7 @vue/composition-api exceljs
# Vue 2.6.x 还需 main.js: Vue.use(require('@vue/composition-api').default)

# 只解析 / 读数据 / 导出(不渲染 UI)
npm i ooxml-excel-editor exceljs

# echarts 可选:仅渲染图表时需要;jspdf 可选:仅导出 PDF 时需要
npm i echarts jspdf
# hyperformula 可选:仅开启编辑 + 公式重算(recalc)时需要
npm i hyperformula
```

四个入口:

| import | 内容 | 需要的 peer | 体积 (gzip) |
|---|---|---|---|
| `ooxml-excel-editor` | **Vue 3** 组件 `<ExcelViewer>` (参考实现 Standard) | `vue@3` + `exceljs` | ~19 KB + 共享 chunks |
| `ooxml-excel-editor/react` | **React** 组件 `<ExcelViewer>` (1:1 复刻 Vue 3) | `react` + `react-dom` + `exceljs` | ~11 KB + 共享 chunks |
| `ooxml-excel-editor/vue2` | **Vue 2.6 / 2.7+** 组件 `<ExcelViewer>` (1:1 复刻 Vue 3) | `vue@2.6+` + `@vue/composition-api` + `exceljs` | ~124 KB (内嵌 core) |
| `ooxml-excel-editor/core` | 框架无关引擎(解析/渲染/控制器/导出/读数据) | `exceljs` | ~1 KB + 共享 chunks |

`exceljs` 必需;`vue` / `react` / `vue@2` 按框架三选一(均为可选 peer);`echarts` / `jspdf` / `hyperformula` 为**可选** peer —— 未装分别只影响"图表渲染""PDF 导出""公式重算",其余正常,且**绝不打包进你的产物**(运行时才动态加载)。

> **三壳 UI 1:1**: Vue 3 SFC 是参考实现 (Standard), Vue 2 / React 1:1 复刻视觉与交互 (工具栏 SVG 图标 / 下拉子菜单 / 公式栏 / 状态栏 / dialog / 浮层 / 演示 demo 全部对齐). 详见 [docs/Vue2.md](./docs/Vue2.md) 跟 Vue 3 的差异速查 + [CLAUDE.md](./CLAUDE.md) 第 7 中心原则。

> ⚠️ **公式重算的许可证**:默认公式引擎是 [HyperFormula](https://hyperformula.handsontable.com/),**GPL-3.0 / 商业 双授权**。本组件以 `licenseKey: 'gpl-v3'` 调用(适合开源/GPL 场景)。**商业闭源项目**请改用 `formulaEngine` prop 注入你自己持有商业 license 的引擎(或自研引擎),只需实现 `FormulaEngine` 接口即可。不开启 `recalc` 时完全不加载 hyperformula,无许可证负担。

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
| `cellReadOnly` | `(cell, pos) => boolean` | 按格只读判定(编辑时) |
| `readOnlyRanges` | `MergeRange[]` | 只读区域(命中即只读,黑名单) |
| `editableTargets` | `EditableTarget \| EditableTarget[]` | **可编辑白名单**(2026-06-08)— 设了就是白名单语义:默认只读,**只**命中**任一** target 的格可编辑。4 种 target 形状自动识别:`{row,col}` 单格 / `{row}` 整行 / `{col}` 整列 / `MergeRange` 矩形。单值或数组都行,允许**不相邻**多 target。`undefined` (不传) = 不启用白名单 = 老行为(默认全可编辑);`[]` (显式空) = 全只读。与 `readOnlyRanges` / `cellReadOnly` 叠加 — 白名单命中后仍可被二次"黑"掉。运行时改:命令式 `viewer.setEditableTargets(targets)` |
| `strictDimensions` | `boolean` | **严格尺寸闸门**(Phase B, 2026-06-08)— 默认 `false`:`setColumnWidth` / `setRowHeight` / `autoFit` 仅受全局 `editable` 控制。设 `true` + 启用了 `editableTargets` → 该列/行至少有 1 格在白名单内才能改尺寸,否则 skip + emit `permission-denied` (`reason='dimension'`)。 |
| `readOnlyCellStyle` | `boolean \| CellStyleOverride \| CellStyleFn` | **只读视觉钩子**(Phase C, 2026-06-08)— 默认 `false` 无视觉差异(老行为);`true` 套内置浅灰底 `#f5f7fa`;对象 = 固定样式给所有只读格;函数 = 按格自定义。仅在该格 `editable=false` 时套用,跟 `editableTargets` 配合一眼看出哪些格可编辑。**鼠标光标**: 编辑模式下悬停只读格自动变 `not-allowed`(内置,不可关)。 |
| `editor` | `EditorResolver` | 自定义单元格编辑器工厂(返回任意 DOM) |
| `recalc` | `boolean` | 公式重算(默认 `false`;需 `editable`) |
| `formulaEngine` | `FormulaEngineFactory` | 自定义/自研公式引擎(默认 HyperFormula) |
| `cellImageFit` | `'fill' \| 'contain' \| 'cover'` | WPS 单元格内嵌图贴合方式(默认 `contain` 等比,与 WPS 渲染一致) |
| `imageLightbox` | `boolean` | 图片点击放大灯箱(默认 `true`;只读单击图放大、编辑右键「查看大图」) |

> 编辑相关 props 详见下方 [编辑](#编辑可选默认只读) 章节。

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

| 导出 | 说明 |
|---|---|
| `ExcelViewer` | 预览/编辑组件 |
| `parseWorkbook(buffer)` | `ArrayBuffer → Promise<WorkbookModel>`(优先 Web Worker) |
| `loadArrayBuffer(src)` | 多种输入归一化为 `ArrayBuffer` |
| `default` | Vue 插件(`app.use`) |
| 类型 | `WorkbookModel` / `SheetModel` / `CellModel` / `CellStyle` / `MergeRange` / `ConditionalRule` / `ChartSpec` / `ImageAnchor` / `CssColor` / `ExcelSource` |

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
| 粘贴 | `pasteText(tsv,at?)` · `pasteRichHtml(html,at?)`(Excel/WPS 复制 → 字体/颜色/填充/边框/对齐/合并 + 单次撤销) · `pasteImageBlob(blob,at?)`(单图/拖文件落格) |
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
| 导出 | `exportXlsx/downloadXlsx` · `exportJson/downloadJson` · `exportCsv/downloadCsv`(见 [导出](#导出--打印)) |

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

### 公式重算(可换引擎)

开 `:recalc` 后,编辑公式格或被公式引用的格 → 依赖格**自动级联重算**,每个变动都发 `cell-change`(`source: 'api'|'ui'|'undo'|'redo'`)。默认引擎 [HyperFormula](https://hyperformula.handsontable.com/)(可选 peer,`npm i hyperformula`;**GPL-3.0/商业双授权**,商业项目用 `:formula-engine` 注入持牌/自研引擎,实现 `FormulaEngine` 接口即可)。`isRecalcReady()` 查引擎是否就绪(异步懒加载)。

### 事件 = 前后完整快照

每次编辑都以 **`cell-change`** 通知:`{ before, after, source }`,`before`/`after` 是 `CellSnapshot` —— 不只 raw,还含**计算值 computed、显示文本 text、整个底层 `CellModel` + 解析后 `style`**。事件流**和**查询 API(`getCellSnapshot`)同一份底层结构 → JSON/CSV/XLSX 导出都复用它,无需按格式各写一遍解析。另有 `dim-change`/`image-change`/`struct-change`/`dirty-change`(见上方事件表)。

### v1 已知限制

- 增删行列**会自动重写公式引用**(`=A5` 上方插一行 → `=A6`,删被引用行 → `#REF!`,含跨表 `Sheet1!A5`)。
- 写回 .xlsx 默认从模型重建,丢 VBA/工作表保护/复杂 DrawingML 等;需要更高保真可用 `exportXlsx({ fidelity: 'overlay' })` **重载原件叠加编辑**(见 [导出保真边界](#导出--打印))。

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
`load(src)` / `getWorkbook()` / `getActiveSheet()` / `setActiveSheet(i)` / `getSelection()` / `setSelection(range)` / `rectOf(row,col)` / `rectOfRange(range)` / `redraw()`,以及下面的导出方法;**编辑命令式 API**(`editCell`/`setStyle`/`insertRows`/`undo`/`exportXlsx`…)见 [编辑](#编辑可选默认只读)。

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

- **`rebuild`(默认)** —— **从编辑后模型完整重建**:遍历 cells/公式/样式(字体/填充/边框/对齐/数字格式)/合并/行高列宽/冻结/图片 重组成 ExcelJS 工作簿。干净、所见即所得,但**丢失**原件里我们不建模的部分(条件格式、数据验证、VBA 宏、工作表保护、复杂 DrawingML/图表 等)。图片导出区分 oneCell/twoCell 锚点 + 子格 EMU 偏移。
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
<ExcelViewer :src="file" />                                           <!-- 默认: find + filter -->
<ExcelViewer :toolbar="['find','filter','separator','zoom','export']" /> <!-- 控制项/顺序/分隔 -->
<ExcelViewer :toolbar="false" />                                      <!-- 隐藏整条 -->
```
- **内置 id**:`find`(查找)、`filter`(切换自动筛选 —— 文件没设也能点出下拉)、`clear-filter`(清除筛选,无筛选时禁用)、`copy`(复制选区)、`wrap-text`(自动换行 toggle,WPS 风格,需 `editable`)、`image-tools`(图片工具 ▾:选区/整表/整列 浮动 ⇄ 嵌入互转,需 `editable`)、`template`(模板 ▾:仅 JSON / 模型数据源下生效;导入 .xlsx 当样式捐赠者;xlsx 数据源下禁用)、`freeze`(冻结/取消)、`zoom`(缩放下拉)、`export`(导出/打印下拉)、`'separator'`/`'|'`(分隔线);`sort` 规划中。
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

## 浏览器支持

现代浏览器(Chrome/Edge 80+、Safari 15+、Firefox 114+,需支持 Canvas / ResizeObserver)。

> **解析线程**:发布的组件库在**主线程**解析(`exceljs` 为 peer 依赖,不重复打包)。本仓库的 demo/dev 额外启用了 **Web Worker** 解析(大文件不卡 UI)。如果你的应用要处理很大的文件,可直接用导出的 `parseWorkbook` 包进你自己的 Worker。

## 范围边界

- **编辑**已支持(默认只读;开 `editable`,见 [编辑](#编辑可选默认只读))。下列为暂不覆盖项:
- 增删行列**不自动重写公式引用**;写回 .xlsx 丢 VBA/工作表保护/复杂 DrawingML
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

> **e2e 说明**:`npm run test:e2e` 用 Playwright 起 dev 服务 + 无头 Chromium,加载示例 → 渲染 → 导出 PNG/位图PDF/矢量PDF,校验产物(PNG 魔数、`%PDF`、矢量 PDF 的文字操作符数量多于位图)。覆盖 node 单测做不到的真实 canvas/jsPDF 绘制。首次需 `npx playwright install chromium` 下载浏览器(本仓库 `@playwright/test` 固定 `1.58.0` 对应 chromium-1208)。Vue demo 在 `/`、React demo 在 `/react.html`。

## 文档 / 二开

- [ARCHITECTURE.md](./ARCHITECTURE.md) —— 包/入口、core 分层、数据流、`ViewerController` 桥接、"加功能改哪"
- [CONTRIBUTING.md](./CONTRIBUTING.md) —— 本地跑通、改动流程、不可破坏的硬约束
- [CHANGELOG.md](./CHANGELOG.md) / [RELEASING.md](./RELEASING.md) —— 变更记录 / 发布清单
- [docs/编辑权限与只读边界.md](./docs/编辑权限与只读边界.md) —— **EditableTarget 白名单 / DimTarget 尺寸多形态 / readOnlyCellStyle 视觉钩子 / permission-denied 事件** 体系化说明(1.2.0)
- [docs/Vue2.md](./docs/Vue2.md) —— **Vue 2 兼容子入口**完整文档(1.3.0;`ooxml-excel-editor/vue2` 跟 Vue 3 / React 壳 ~100% 功能对齐)

> **React props/events** 与 Vue 对齐(事件用 camelCase 回调:`onRendered`/`onError`/`onCellClick`/`onSelectionChange`/`onSheetChange`/`onHyperlinkClick`),命令式句柄 `ExcelViewerHandle` 与 Vue 组件 ref 同名方法一致。上面「扩展 API」中的**插件 `definePlugin`** 目前服务 Vue 壳;React 壳已可用全部 props/命令式 API/事件,插件 overlay 跨框架化在路线图中。

## License

MIT

---

<a name="english"></a>
## English

A **Vue 3 + React high-fidelity `.xlsx` preview & edit component** with a from-scratch parser and canvas renderer. Renders cells, number formats, merges, conditional formatting, images, charts (via ECharts), sparklines, comments, data validation, frozen panes, and supports selection / copy / hyperlinks. **Read-only by default**; set `editable` to enable editing — cell values / styles / column-row sizes / floating images / insert-delete rows-cols, with undo-redo, before/after full-snapshot events, dirty tracking + reset-to-original, swappable formula recalc engine, custom cell editors, and **export back to .xlsx / JSON / CSV**. Parsing runs in a Web Worker (with main-thread fallback). `vue` / `react` / `exceljs` are peer dependencies; `echarts` / `jspdf` / `hyperformula` are optional peers (charts / PDF / formula recalc).

```bash
npm i ooxml-excel-editor vue exceljs
```

```ts
import { ExcelViewer } from 'ooxml-excel-editor'
import 'ooxml-excel-editor/style.css'
```

See the API table above. MIT licensed.
