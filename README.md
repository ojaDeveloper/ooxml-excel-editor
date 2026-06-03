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
- 🖱 **交互**:单元格选区(合并感知)、拖选、公式栏、状态栏(计数/求和/均值/最值)、超链接可点、裁切文本悬停看全文、Ctrl+C 复制为 TSV
- ⚡ **按需加载**(无图表文件不下载 echarts)、**友好错误兜底**(损坏/加密/旧 .xls)、解析失败自动给出可读提示

> 预览不需要公式引擎 —— .xlsx 缓存了公式结果,直接显示。详见 [EXCEL还原难点.md](./EXCEL还原难点.md)。

## 安装

```bash
npm i ooxml-excel-preview vue exceljs
# echarts 可选:仅当要渲染图表时才需要
npm i echarts
```

`vue` / `exceljs` / `echarts` 是 **peerDependencies**(由宿主项目提供,组件本身不重复打包)。`echarts` 为可选 —— 未安装时,含图表的文件会显示占位提示,其余正常。

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

## 浏览器支持

现代浏览器(Chrome/Edge 80+、Safari 15+、Firefox 114+,需支持 Canvas / ResizeObserver)。

> **解析线程**:发布的组件库在**主线程**解析(`exceljs` 为 peer 依赖,不重复打包)。本仓库的 demo/dev 额外启用了 **Web Worker** 解析(大文件不卡 UI)。如果你的应用要处理很大的文件,可直接用导出的 `parseWorkbook` 包进你自己的 Worker。

## 范围边界(第一版不做)

- 编辑 / 公式重算(只显示缓存结果)
- 透视表:**数据按普通单元格显示**,但无字段按钮/下拉等透视专属 UI
- SmartArt;形状仅支持 rect/roundRect/ellipse + 文本(复杂自定义几何按矩形近似)
- `.xls`(旧 BIFF 二进制)/ 加密文件(给出友好提示)
- 图表为 ECharts 近似,非像素级一致

## 开发

```bash
npm install
npm run dev            # 本地预览(demo)
node scripts/gen-sample.mjs   # 生成 public/sample.xlsx 示例
npm run test           # 单元 + 端到端测试
npm run typecheck      # 类型检查
npm run build          # 构建组件库(dist/)
npm run build:demo     # 构建 demo 站点
```

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
