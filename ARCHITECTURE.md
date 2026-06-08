# 架构(ARCHITECTURE)

> 给**二开者**看的地图:包怎么分、core 怎么分层、数据怎么流、加功能改哪。
> 调用方文档看 [README](./README.md);贡献流程看 [CONTRIBUTING](./CONTRIBUTING.md);三壳 1:1 复刻原则看 [CLAUDE.md](./CLAUDE.md) 第 7 中心原则。

## 一句话

**一个框架无关的 TS 引擎(`core`)+ 三个薄壳(Vue 3 / Vue 2 / React)**。引擎做全部解析、渲染、交互、查找、筛选、导出;壳只负责"把引擎接到框架的生命周期 + 渲染 chrome(工具栏/公式栏/状态栏/标签/弹层)"。**Vue 3 是参考实现 (Standard), Vue 2 / React 1:1 复刻**(UI / props / events / 演示 demo 视觉与交互一致)。三壳**逻辑同构、共用 ~100% 引擎**。

## 包 / 入口

单包,四个子入口,**共享同一份 `dist/core.js`**(引擎只打一份)+ Vue 2 独立 build:

| import 路径 | 产物 | 内容 | peer |
|---|---|---|---|
| `ooxml-excel-editor` | `dist/index.js` | Vue 3 壳 `<ExcelViewer>` | `vue@3` + `exceljs` |
| `ooxml-excel-editor/react` | `dist/react.js` | React 壳 `<ExcelViewer>` | `react`+`react-dom`+`exceljs` |
| `ooxml-excel-editor/vue2` | `dist/vue2.js` | Vue 2.7 壳 `<ExcelViewer>` (1.3.0) | `vue@2.7` + `exceljs` |
| `ooxml-excel-editor/core` | `dist/core.js` | 框架无关引擎 | `exceljs` |

`index.js` / `react.js` 都 `import './core.js'`(共享 chunk)。`vue@3`/`react`/`react-dom`/`exceljs`/`echarts`/`jspdf` 全 external,不打进产物。

**Vue 2 入口** 用独立 build pass (`vite build --mode lib-vue2`), 因为 vue@3 + vue@2 同时存在导致 `@vitejs/plugin-vue` / `plugin-vue2` 的 `@vue/compiler-sfc` 解析冲突 — 单 entry 隔离才能跑通。代价: `vue2.js` 396 KB 含内嵌 core, 不像主入口那样共享 `chunks/`。架构上仍跟其他壳同构 — `ViewerController + hooks` 模式, Vue 2.7 内置 Composition API + render function (`.ts`, 不走 SFC)。详见 [docs/Vue2.md](./docs/Vue2.md).

生态变大时可平滑拆成真正的 workspace 三包,**无需改源码结构**(壳早已只依赖 `core/` 的公共导出)。

## core 分层(`src/core/`)

```
index.ts   框架无关公共入口(引擎 + 解析 + 数据 + 类型)
parser/    .xlsx 解压 + ExcelJS 适配 + 原始 XML 薄层(theme/drawings/charts/sparklines/pageSetup
           + WPS 内嵌图 cell-image-parser(DISPIMG)+ row-meta-parser(customHeight))
loader.ts  多种输入(File/Blob/ArrayBuffer/URL)归一成 ArrayBuffer
finalize.ts 文件头探测 / 图片 blob 化 / 友好错误
model/     中间模型 types(WorkbookModel/SheetModel/CellModel)+ data-access(读)+ mutations(写)+ clone(快照)
edit/      EditController(命令栈)+ commands + clipboard-html(富粘贴解析)+ context-menu + default-editor + editor-host
layout/    grid-metrics(行列几何 + 虚拟外推)/ merges / freeze / autofit / viewport
format/    number-format(数字/日期格式 mini 引擎)+ color(toHex6 等)/ date-serial
render/    canvas-renderer(普通类,画一帧;含 DISPIMG 内嵌图绘制)+ conditional/fills/borders/text/autofilter/theme
overlay/   anchor(锚点几何)/ chart-mapper(图表→echarts option)/ echarts-loader
export/    raster/composite/paginate/pdf/print/vector-pdf + xlsx-writer + wps-cellimages(DISPIMG 回注)+ WorkbookExporter
           + abort.ts(yieldToEvent + checkAborted):所有耗时导出统一接 onProgress + AbortSignal,
             串行 + 每阶段 emit + 让出 UI(防假死)+ 可中断
loader-json.ts  JSON → WorkbookModel(P3:三种 shape 自动识别 + 类型推断;绕过 parser)
template/  fill.ts  占位符 {{key}} + 锚点表(startCell + rows)注入,渲染前预处理(不入命令栈)
viewer/    OverlayManager(图片/图表/形状 DOM 叠加层)+ LightboxHost(图片放大)
           ViewerController(总编排:renderer 生命周期 + view 状态 + 选区 + 鼠标/键盘 + 编辑 + 查找 + 筛选 + 导出)
```

**铁律**:`src/core/**` 不得 `import` 任何框架(`vue`/`react`)。构建后 `dist/core.js` 也不得出现框架 import(CI 可加 grep 守门)。

## 壳里的 UI 区域(命名 + 替换方式)

```
┌─ <ExcelViewer> ─────────────────────────────────────────┐
│ ① Header 顶栏          ViewerToolbar           #header slot
│ ② ActionToolbar 操作栏 .action-toolbar         :toolbar 配置 + 插件
│ ③ FormulaBar 公式栏 Fx .formula-bar            CSS / 命令式 API
│ ④ RenderArea 网格      .render-area + canvas   theme / cellStyle / overlay slot
│    ├ OverlayManager    图片/图表/插件 DOM 叠加
│    ├ ContextMenu       body 级右键菜单(editable)
│    ├ LightboxHost      body 级图片放大 + 下载
│    └ CellEditorHost    自定义编辑器挂载点(:editor)
│ ⑤ SheetTabs 表标签     .sheet-tabs
└─────────────────────────────────────────────────────────┘
```

完整速查表 + 文件路径 + 替换方式见 [README「UI 区域速查」](./README.md#ui-区域速查给调用方--二开者)。
**注**:demo 跑 `npm run dev` 看到的绿色 `.app-bar` 是 demo 框架栏(src/App.vue / src/react-demo/main.tsx),不属于 `<ExcelViewer>` 自带 chrome,用户 import 不会出现。

## 数据流

```
文件/URL → loader → parser → 中间模型 WorkbookModel
                                     │
                          ViewerController.rebuild()
                                     │
                         CanvasRenderer.render(view) → <canvas>
                         OverlayManager.position()    → 图片/图表/形状 DOM
```

中间模型与 ExcelJS / 原始 XML 形状**完全解耦** —— 换解析库或换渲染后端,只要产出/消费这套 `WorkbookModel` 即可。读数据 API(`getSheetData`/`sheetToJSON`/…)直接吃模型,不经渲染。

## ViewerController —— 壳与引擎的唯一桥

壳在挂载后 `new ViewerController(els, hooks)`:

- **els**:`{ canvas, renderArea, scroller, spacer, overlays:{main,frow,fcol,corner} }` —— 壳提供的 DOM 节点。
- **hooks**:引擎回调壳的一组函数,驱动框架重渲:
  `onRenderer`(renderer 重建)、`onRenderTick`(每帧绘制后)、`onSelectionChange`、`onCellClick`、`onCellDblClick`、`onHyperlink`、`onTooltip`、`onFindChange`、`onFilterChange`。

控制器对外暴露命令式方法:`rebuild / render / setScroll / setZoom / 选区(getSelection/setSelectionRange/selectCell/…) / 交互(onMouseDown/Move/Up/DblClick/KeyDown) / 查找(setFindQuery/findNext/…) / 筛选(toggleAutoFilter/applyFilterSelection/…) / 导出(exportImage/downloadPdf/…) / rectOf`。

壳侧桥接模式:
- **Vue**:hooks 里 `xxxVersion.value++`(`ref` 计数器)→ 计算属性重算;`onMounted` 建控制器,`onBeforeUnmount` `dispose()`。
- **React**:hooks 里 `force()`(`useReducer` 派发)→ 重渲;**控制器创建 + rebuild 用 `useLayoutEffect`**(同步绘制,避免晚到的 passive rebuild 清掉刚设的交互态)。

导出编排单独抽在 `export/WorkbookExporter`,靠一个 `ExporterHost` 取数器(工作簿/活动表/复用 renderer/渲染配置/文件名)与控制器解耦,壳不感知。

## 加功能改哪(决策树)

- **新渲染能力 / 新交互 / 新查筛逻辑** → 做进 `ViewerController`(或它依赖的 `core/render`、`core/layout`)。**三壳都接上**,各自补 e2e。
- **新解析字段** → `parser/` + `model/types`(+ 若要读出来,`model/data-access`)。
- **新导出格式 / 分页规则** → `export/`(`WorkbookExporter` 编排,`pdf`/`vector-pdf`/`print` 实现)。
- **纯外观 / chrome(按钮、弹层样式)** → **必须 Vue 3 先落地**(`src/components/ExcelViewer.vue` 或子 SFC,Standard),然后 1:1 复刻到 Vue 2 (`src/vue2/ExcelViewer.ts` render function) 和 React (`src/react/ExcelViewer.tsx`)。可跨框架共享的纯 TS / CSS 抽到 `src/components/toolbar-icons.ts` / `src/demo-shared/`。详见 [CLAUDE.md](./CLAUDE.md) 第 7 中心原则。
- **demo 入口演示功能 / 按钮** → Vue 3 demo (`src/App.vue`) 先落地, 共享 CSS 放 `src/demo-shared/demo-bar.css`, Vue 2 (`vue2-demo/main.ts`) + React (`src/react-demo/main.tsx`) 1:1 复刻按钮列表 / 顺序 / 文案 / title。
- **可配置扩展点**(主题/钩子/插件/事件/overlay)→ 见 README「扩展 API」;新增能力优先做成可配置/可覆盖,而非写死。

## 三壳 UI 1:1 复刻 (★ 第 7 中心原则)

Vue 3 SFC 是参考实现 (Standard), Vue 2 / React 复刻. 任何 UI 类改动顺序:

1. **Vue 3 先落地** (`src/components/*.vue`):template + scoped CSS + setup
2. **抽共享** (可跨框架复用):
   - 纯 TS 模板/图标/类型 → `src/components/toolbar-icons.ts` / `toolbar-types.ts` / `export-types.ts`
   - 全局 CSS (三 demo 共享) → `src/demo-shared/demo-bar.css`
3. **React 复刻** (`src/react/ExcelViewer.tsx`):JSX 同结构 + className 同款 (或 `rxl-` 前缀但视觉对齐)
4. **Vue 2 复刻** (`src/vue2/ExcelViewer.ts`):render function (h() + VNode), **每个 VNode 必须 key**, controller 持有的 DOM 在 onMounted createElement + appendChild
5. **三 demo 也要 1:1** (顶栏按钮 / 顺序 / 文案)
6. **验证** typecheck + 单测 + Playwright 三 demo 截图对比 + build 全绿

详见 [CLAUDE.md](./CLAUDE.md) 「UI 1:1 复刻工作流」段。

## 测试结构

- **单测(vitest,node 环境)**:`src/**/__tests__/`,覆盖 parser/layout/format/render/export/data-access/edit/plugin —— 纯逻辑回归网(316)。
- **e2e(Playwright,真 Chromium)**:`e2e/*.e2e.ts`,覆盖 canvas 真绘制 / jsPDF 下载 / 查找 / 筛选 / 工具栏 / 数据 API,Vue 3 走 `/`、React 走 `/react.html`(60)。Vue 2 e2e 后续补 (multi-entry Playwright config)。

改 core 行为务必三壳都跑绿(见 CONTRIBUTING)。
