# ooxml-excel-editor

> Vue 3 + React 高保真 **.xlsx 预览 / 编辑组件** —— Canvas 渲染,**默认只读预览,可选开启编辑**。从零实现解析与渲染,尽量还原微软 Excel 打开工作簿的观感。

[English](#english) · 中文

## ⚡ 快速开始

**装**(按框架二选一):

```bash
npm i ooxml-excel-editor vue exceljs                 # Vue
npm i ooxml-excel-editor react react-dom exceljs     # React
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

默认**只读预览**;想编辑加 `:editable="true"`(React 同名 `editable`)。React 写法、props/事件表、编辑 / 导出 API 见下文对应章节。

> 纯使用者只需读 **安装 / 使用 / API / 编辑 / 导出** 几节即可接入,无需看源码;类型随包发 `.d.ts`(IDE 自动补全)。「扩展 API / 插件 / 开发」是进阶,可跳过。

## 特性

- 📊 **Canvas 高保真渲染**:DPR 高清、虚拟滚动(万行流畅)、冻结窗格四象限
- 🔢 **自写数字格式引擎**:千分位/货币/百分比/科学计数/分数、四段格式(正;负;零;文本)、`[Red]` 颜色、`[>=100]` 条件段、中文日期 `yyyy"年"`、`[h]:mm` 经过时间
- 🗓 **日期序列号**:含 Excel 1900 闰年 bug、1904 系统
- 🎨 **主题色 + tint**、indexed 调色板、合并单元格、边框(细/粗/虚/双线)、填充(纯色/图案/渐变)
- 🌈 **条件格式**:色阶 / 数据条 / 图标集 / cellIs / top10
- 🖼 **图片 + 图表**(DrawingML → ECharts 近似还原)、**形状/文本框**、**迷你图**(sparklines)、**批注**、**数据验证**下拉、**自动筛选**样式
- 📝 **文本溢出**到相邻空格、**自动行高**
- 🖱 **交互**:单元格选区(合并感知)、拖选、公式栏、状态栏(计数/求和/均值/最值)、超链接可点、裁切文本悬停看全文、Ctrl+C 复制为 TSV、**Ctrl+F 查找**(高亮 + 上/下定位 + 计数 + 区分大小写/全字匹配)、**自动筛选**(点下拉真能筛:去重值多选 + 搜值 + 清除)
- 🖨 **导出 / 打印**:整表/选区/多表导出 **PNG/JPEG**、**PDF**(位图 + **矢量·文字可选可搜**两种)、**系统打印**(可另存 PDF);默认还原原生 `pageSetup`(纸张/方向/页边距/缩放/打印区域/**打印标题行列每页重复**);宽表**横向跨页**(页矩阵);`beforeRenderPage` 注入页眉/页脚/水印、`configureDoc` 注册字体;内置「导出设置」对话框
- ⚡ **按需加载**(无图表文件不下载 echarts、不导出 PDF 不下载 jspdf)、**友好错误兜底**(损坏/加密/旧 .xls)、解析失败自动给出可读提示

- 📤 **数据读取 API**:不必自己再解析 —— `getCellText`/`getSheetData`/`sheetToJSON`/`getRangeData`(独立函数 + 组件 ref 方法),值/显示文本可选,合并/日期/数字格式都处理好
- ✏️ **编辑(可选,默认只读)**:开 `editable` 即可编辑 —— 单元格值 / 样式(粗体/对齐/填充)/ 列宽行高 / 浮动图片(拖拽移改)/ 增删行列;**撤销重做**(Ctrl+Z/Y)、**前后完整快照事件**、**脏状态 + 一键还原原件**;可换**公式引擎**自动重算依赖格;可注入**自定义编辑器**(下拉/日期/图片选择器);**导出回 .xlsx / JSON / CSV**(所见即所得)。见 [编辑](#编辑可选默认只读)

> 纯预览不需要公式引擎 —— .xlsx 缓存了公式结果,直接显示;仅开启**编辑 + 重算**时才用(可选 `hyperformula`)。详见 [EXCEL还原难点.md](./EXCEL还原难点.md)。

## 安装

一个包,三个子入口 —— **框架无关的 core 引擎被 Vue / React 两个薄壳共享**(`dist/core.js` 只打一份)。按你的框架装对应 peer:

```bash
# Vue 项目
npm i ooxml-excel-editor vue exceljs

# React 项目
npm i ooxml-excel-editor react react-dom exceljs

# 只解析 / 读数据 / 导出(不渲染 UI)
npm i ooxml-excel-editor exceljs

# echarts 可选:仅渲染图表时需要;jspdf 可选:仅导出 PDF 时需要
npm i echarts jspdf
# hyperformula 可选:仅开启编辑 + 公式重算(recalc)时需要
npm i hyperformula
```

三个入口:

| import | 内容 | 需要的 peer |
|---|---|---|
| `ooxml-excel-editor` | Vue 3 组件 `<ExcelViewer>` | `vue` + `exceljs` |
| `ooxml-excel-editor/react` | React 组件 `<ExcelViewer>` | `react` + `react-dom` + `exceljs` |
| `ooxml-excel-editor/core` | 框架无关引擎(解析/渲染/控制器/导出/读数据) | `exceljs` |

`exceljs` 必需;`vue` / `react` / `react-dom` 按框架二选一(均为可选 peer);`echarts` / `jspdf` / `hyperformula` 为**可选** peer —— 未装分别只影响"图表渲染""PDF 导出""公式重算",其余正常,且**绝不打包进你的产物**(运行时才动态加载)。

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

## API

### `<ExcelViewer>`

| Prop | 类型 | 说明 |
|---|---|---|
| `src` | `File \| Blob \| ArrayBuffer \| Uint8Array \| string(URL)` | 要预览的 .xlsx 数据源 |
| `fileName` | `string` | 标题栏显示的文件名(可选) |
| `editable` | `boolean` | 开启编辑(默认 `false` = 只读,行为与历史一致) |
| `cellReadOnly` | `(cell, pos) => boolean` | 按格只读判定(编辑时) |
| `readOnlyRanges` | `MergeRange[]` | 只读区域(命中即只读) |
| `editor` | `EditorResolver` | 自定义单元格编辑器工厂(返回任意 DOM) |
| `recalc` | `boolean` | 公式重算(默认 `false`;需 `editable`) |
| `formulaEngine` | `FormulaEngineFactory` | 自定义/自研公式引擎(默认 HyperFormula) |

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

开编辑后:**双击 / F2 / 直接打字**进格编辑(Enter 提交下移、Tab 右移、Esc 取消),**Ctrl+Z / Ctrl+Y** 撤销重做,拖列头/行头边界改宽高,拖浮动图片移动。只读格 / 区域不进编辑。

### 命令式编辑 API(模板 ref / 插件 `viewer`)

| 类别 | 方法 |
|---|---|
| 值 | `editCell(row,col,value)` · `editRange(range,values[][])` · `clearRange(range)` |
| 样式 | `setStyle(range, patch)`(`patch` = `CellStyleOverride`:font/fill/borders/对齐/numFmt) |
| 列宽行高 | `setColumnWidth(col,px)` · `setRowHeight(row,px)` |
| 行列结构 | `insertRows(at,count?)` · `deleteRows(at,count?)` · `insertCols(at,count?)` · `deleteCols(at,count?)` |
| 图片 | `getImages()` · `addImage(anchor)` · `removeImage(i)` · `moveImage(i,dxPx,dyPx)` · `resizeImage(i,wPx,hPx)` |
| 撤销/进编辑 | `undo()` · `redo()` · `canUndo()` · `canRedo()` · `beginEdit(row,col)` · `cancelEdit()` · `isEditing()` · `getEditingCell()` |
| 查询/状态 | `getCellSnapshot(row,col)` · `isDirty()` · `resetToOriginal()` · `isRecalcReady()` |
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

### 操作工具栏(可配置 / 可插件 / 响应式)
顶栏(文件名/导出/缩放)下方有一行**操作工具栏**,内置 `find`/`filter` 按钮默认显示。用 `:toolbar` 配置:
```vue
<ExcelViewer :src="file" />                                           <!-- 默认: find + filter -->
<ExcelViewer :toolbar="['find','filter','separator','zoom','export']" /> <!-- 控制项/顺序/分隔 -->
<ExcelViewer :toolbar="false" />                                      <!-- 隐藏整条 -->
```
- **内置 id**:`find`(查找)、`filter`(切换自动筛选 —— 文件没设也能点出下拉)、`clear-filter`(清除筛选,无筛选时禁用)、`copy`(复制选区)、`freeze`(冻结/取消)、`zoom`(缩放下拉)、`export`(导出/打印下拉)、`'separator'`/`'|'`(分隔线);`sort` 规划中。
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
