# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与 [语义化版本](https://semver.org/lang/zh-CN/)。
尚未发布到 npm —— 首个 `npm publish` 即 `0.1.0`。

## [Unreleased]

### 新增
- **React 壳**:`ooxml-excel-preview/react` 导出 `<ExcelViewer>`(`forwardRef` + 命令式 `ExcelViewerHandle`)与 `useExcelDocument`,与 Vue 壳**共用 ~100% core 引擎**。
- **框架无关 core 入口**:`ooxml-excel-preview/core` 暴露引擎(`ViewerController` / `CanvasRenderer` / `WorkbookExporter` / `OverlayManager`)+ 解析 + 读数据 + 类型,零框架依赖。
- **多入口构建**:产物拆为 `dist/core.js`(引擎)+ `dist/index.js`(Vue 壳)+ `dist/react.js`(React 壳),后两者共享同一份 `core.js`;各自 `.d.ts`。
- React demo(`/react.html`)+ React 真浏览器 e2e(渲染 / 选区 / 查找 / 数据 API / PNG 导出)。

### 变更(重构,行为零回归)
- 把 `ExcelViewer.vue` 的非框架编排逐步下沉到框架无关 `src/core/viewer/`:
  - `OverlayManager`(图片/图表/形状叠加层)
  - `ViewerController`:渲染引擎(renderer/view/scroll/zoom/spacer)、选区模型 + 鼠标/键盘交互、查找引擎、自动筛选引擎
  - `WorkbookExporter`(`src/core/export/exporter.ts`):导出/打印编排,靠 `ExporterHost` 取数器与壳解耦
- `ExcelViewer.vue` 收薄为薄壳:props/插件桥接 + chrome + 经 hooks 桥接控制器响应式。
- `package.json`:`exports` 增 `./react`、`./core`;`vue`/`react`/`react-dom` 改为可选 peer(按框架二选一)。

### 修复
- React 壳:控制器创建与 rebuild 改用 `useLayoutEffect`,避免晚到的 passive rebuild 清掉刚设置的选区/交互态。
- 库构建:`worker-client` stub 别名兼容 React 壳的 `@/composables/worker-client` 引入,避免误把 1.4MB exceljs 打进产物。

### 基线
- 测试:93 单测 + 13 e2e(12 Vue + 1 React)全绿;`dist/core.js` 无 vue/react import;exceljs 仅运行时 `import()` 不打包。

---

首版前的功能积累(解析 / canvas 渲染 / 冻结窗格 / 合并单元格 / 条件格式 / 图表 / 图片形状 / 数字日期格式 / 超链接批注 / 查找 / 自动筛选 / 可配置可插件工具栏 / 图片·位图PDF·矢量PDF·打印导出 / 读数据 API / 主题·钩子·插件扩展点)见 git 历史与 [README](./README.md)。
