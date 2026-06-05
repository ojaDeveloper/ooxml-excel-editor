# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与 [语义化版本](https://semver.org/lang/zh-CN/)。

## [1.0.0] - 2026-06-05

**只读 → 可编辑** —— 在 0.2.0(只读双壳)基础上,把组件升级成**可选编辑器**(默认仍只读、零回归)。
所有编辑能力建在**一份可变内存模型**上,读 / 写 / 事件 / 导出共用同一层;core 始终框架无关、两壳同构。

### 编辑能力(props `editable` 开启)
- **单元格编辑**:双击 / F2 / 打字进编辑,内置文本编辑器 + `editor` 钩子自定义(下拉/日期/图片选择器,返回任意 DOM,框架无关)。
- **命令栈**:撤销/重做(Ctrl+Z/Y),所有写操作(值/样式/宽高/图片/结构)统一进栈。
- **样式编辑**:`setStyle(range, patch)`(font/fill/borders/对齐/数字格式)。
- **列宽行高**:拖拽 + `setColumnWidth/setRowHeight`,入命令栈。
- **图片**:浮动图拖拽移动 + `addImage/removeImage/moveImage/resizeImage`。
- **行列结构**:`insertRows/deleteRows/insertCols/deleteCols`(重键 cells、移合并/宽高/图片,快照逆撤销)。
- **公式重算**:`recalc` 开启 → 依赖格自动级联;默认 HyperFormula(可选 peer,GPL/商业双授权),`formulaEngine` 可换自研/持牌引擎(`FormulaEngine` 接口)。
- **脏状态**:`isDirty()` + `dirty-change` 事件 + `resetToOriginal()`(放弃修改还原原件)。

### 事件 + 查询(底层机制)
- **`cell-change` 携前后完整快照**(`CellSnapshot`:底层 CellModel + 解析 style + raw/computed/text),与 `getCellSnapshot` 查询 API 同一份结构。
- `dim-change` / `image-change` / `struct-change` / `dirty-change` / `edit-start` / `edit-commit`。

### 导出(一份数据层,所见即所得)
- **`.xlsx`**(`exportXlsx/downloadXlsx`):从编辑后模型重建(需 `exceljs`),覆盖值/公式/样式/合并/宽高/冻结/图片。
- **`.json`**(raw 类型值)/ **`.csv`**(格式化显示值,UTF-8 BOM):复用 `getSheetData/getWorkbookJSON`,三格式天然一致。

### 限制
- 增删行列**不自动重写公式引用文本**(缓存值随格移动);写回 .xlsx 丢 VBA/工作表保护/复杂 DrawingML;图片导出忽略部分子格偏移。

### 测试
- 175 单测 + 38 真浏览器 e2e(Vue + React 双覆盖)+ build 全绿;`dist/core.js` 零 vue/react/hyperformula/exceljs 静态 import(重依赖全动态懒加载)。

## [0.2.0] - 2026-06-05

**只读** —— 在 0.1.0(Vue-only v1)基础上加:**Vue + React 双壳共享框架无关 core**、列排序、
跨框架插件、边框还原、多入口分包(core/vue/react)、完整文档。后续「编辑」能力按 semver 推进(0.3.0+)。

### 交互 / 插件
- **列排序**:自动筛选下拉加「升序/降序」,按列重排数据区(整行移动),合并区相交则拒绝。
- **插件跨框架**:`overlay` 钩子从返回 Vue VNode 改为返回 **DOM 节点**,`core/plugin` 不再 import vue(core 彻底框架无关)。**同一份 `definePlugin` 在 Vue 与 React 壳通用**;React 壳新增 `plugins` prop,支持 theme/transformModel/cellStyle/events/overlay/toolbar/setup 全套。

### 边框还原(对齐 Excel/WPS)
- **合并单元格内部不再画网格线**(之前无填充的合并格会透出内部浅灰网格线)。
- **斜线边框(对角线 ↘/↗)**:parser 解析 `diagonal{up,down,style,color}`,canvas 与矢量 PDF 都绘制。
- **相邻共享边按权重取较重者**(hair<…<medium<thick<double):普通格的边框绘制顺序无关、与 Excel/WPS 一致(合并区仍画自身四周)。

### 新增
- **React 壳**:`ooxml-excel-preview/react` 导出 `<ExcelViewer>`(`forwardRef` + 命令式 `ExcelViewerHandle`)与 `useExcelDocument`,与 Vue 壳**共用 ~100% core 引擎**。
- **框架无关 core 入口**:`ooxml-excel-preview/core` 暴露引擎(`ViewerController` / `CanvasRenderer` / `WorkbookExporter` / `OverlayManager` / `PluginOverlayHost`)+ 解析 + 读数据 + 类型,零框架依赖。
- **多入口构建**:产物拆为 `dist/core.js`(引擎)+ `dist/index.js`(Vue 壳)+ `dist/react.js`(React 壳),后两者共享同一份 core 引擎 chunk;各自 `.d.ts`。
- React demo(`/react.html`)+ React 真浏览器 e2e(渲染 / 选区 / 查找 / 数据 API / 导出 / 插件)。

### 变更(重构,行为零回归)
- 把 `ExcelViewer.vue` 的非框架编排逐步下沉到框架无关 `src/core/viewer/`:
  - `OverlayManager`(图片/图表/形状叠加层)+ `PluginOverlayHost`(插件 overlay DOM)
  - `ViewerController`:渲染引擎、选区 + 鼠标/键盘交互、查找、自动筛选、排序、导出编排桥接
  - `WorkbookExporter`(`src/core/export/exporter.ts`):导出/打印编排,靠 `ExporterHost` 与壳解耦
- `ExcelViewer.vue` 收薄为薄壳:props/插件桥接 + chrome + 经 hooks 桥接控制器响应式。
- `package.json`:`exports` 增 `./react`、`./core`;`vue`/`react`/`react-dom` 改为可选 peer(按框架二选一)。

### 修复
- React 壳:控制器创建与 rebuild 改用 `useLayoutEffect`,避免晚到的 passive rebuild 清掉刚设置的交互态。
- 库构建:`worker-client` stub 别名兼容 React 壳的 `@/composables/worker-client` 引入,避免误把 1.4MB exceljs 打进产物。

### 基线
- 测试:**111 单测 + 16 e2e(Vue + React)**全绿;`dist/core.js` 无 vue/react import;exceljs 仅运行时 `import()` 不打包。

## [0.1.0] - 2026-06-03

高保真 .xlsx 预览组件 v1:从零实现解析 + Canvas 渲染(Vue,只读)。冻结窗格 / 合并单元格 /
条件格式 / 图表 / 图片形状 / 数字日期格式 / 超链接批注 / 查找 / 自动筛选 / 可配置可插件工具栏 /
导出(图片·位图PDF·矢量PDF·打印) / 读数据 API / 主题·钩子·插件扩展点。
