# ooxml-excel-editor — 项目开发准则(AI 与贡献者必读)

> Vue3 + React 高保真 .xlsx 预览/编辑组件。从零实现解析与 Canvas 渲染。**默认只读、零回归;开 `editable` 进入编辑**(值/样式/列宽行高/图片/增删行列/公式重算/导出回写 .xlsx·JSON·CSV)。

## ★ 六条中心原则(任何后续开发都要围绕,不得破坏)

1. **好文档** —— 同时服务**调用方**(怎么用)和**二开者**(怎么改/扩展)。改公开 API/扩展点/导出,必须同步更新 README + 相关文档(ARCHITECTURE/CONTRIBUTING/各包 README + CHANGELOG)。README 的 props 表 / 导出表 / 选项表要与代码一致。
2. **可发布** —— 始终保持能 `npm publish` 的状态:exports/main/module/types 对、`.d.ts` 完整、peer 依赖(vue/react/exceljs 必需,echarts/jspdf 可选)**绝不打包进产物**、占位元数据要清。
3. **Vue + React 共存** —— `core` 框架无关,Vue/React 各是薄壳共享同一 `core`。**禁止在 `core` 里 import vue/react**。新功能优先做进 `core`(框架无关),壳只做桥接。
4. **包名清晰** —— 现状: **单包三子入口**(`ooxml-excel-editor` = Vue 壳 / `/react` = React 壳 / `/core` = 框架无关引擎),三者共享同一 `dist/core.js`。后续生态大了可平滑拆成真正的 workspace 三包(`@scope/core`+`/vue`+`/react`);现阶段单包多入口已满足"按框架各取所需",不过度拆包。
5. **扩展点** —— 保留并尊重已有扩展点:`:theme` / `transformModel` / `cellStyle` / 事件 / `overlay` slot / `rectOf` 命令式 API。新增能力优先做成可配置/可覆盖,而非写死。跨框架时把 Vue 特定扩展点(如 `overlay` 返回 VNode)做成框架无关(返回 DOM/描述)。
6. **插件机制** —— `definePlugin` 打包 theme/transformModel/cellStyle/events/overlay/toolbar/setup;多插件按数组合并、组件 props 最后覆盖。改动不得破坏此契约;插件应跨框架可用。

## 架构(framework-agnostic core + 薄壳)

```
src/core(纯 TS,零框架依赖)→ 构建产物 dist/core.js
  index.ts  框架无关公共入口(引擎 + 解析 + 数据 + 类型)
  parser/   ExcelJS 适配 + 原始 XML 薄层(theme/drawings/charts/sparklines/pageSetup)
  model/    中间模型 types + data-access(读数据 API)
  layout/   grid-metrics / merges / freeze / autofit / viewport
  render/   canvas-renderer(普通类) + conditional/fills/borders/text/autofilter/theme
  format/   number-format(数字/日期格式 mini 引擎)
  overlay/  anchor / chart-mapper / echarts-loader
  export/   raster/composite/paginate/pdf/print/vector-pdf + WorkbookExporter(导出编排)
  viewer/   OverlayManager + ViewerController(渲染/选区/交互/查找/筛选/导出 全编排,框架无关)✅
src/components(Vue 壳)→ dist/index.js:ExcelViewer.vue + 子 SFC(toolbar/find/filter/dialog/tabs)
                        薄壳, onMounted 起 ViewerController, hooks 桥接响应式
src/react(React 壳)→ dist/react.js:ExcelViewer.tsx + use-excel-document
                        薄壳, useLayoutEffect 起同一 ViewerController, useReducer force 桥接
```
数据流:文件 → loader → parser → **中间模型(WorkbookModel)** → ViewerController(含 CanvasRenderer)→ canvas。
中间模型与 ExcelJS/XML 形状**完全解耦**。`ViewerController` 是壳与引擎的唯一桥:壳给 DOM 元素 + 一组回调 hooks
(onRenderer/onRenderTick/onSelectionChange/onCellClick/…),控制器回调驱动壳的响应式重渲。**Vue 与 React 壳逻辑同构、共用 ~100% 引擎**。

## 不可破坏的硬约束

- **测试是回归网**:改动后 `npm run typecheck` + `npm test`(单测)+ `npm run test:e2e`(Playwright 真浏览器)+ `npm run build` 必须全绿。当前基线 **188 单测 + 40 e2e(Vue + React 双覆盖)**。
- **core 不依赖框架**:`src/core/**` 不得出现 `from 'vue'` / `'react'`(构建后 `dist/core.js` 也不得 import vue/react/hyperformula/exceljs —— 重依赖全动态懒加载)。
- **两壳同构**:给 `ViewerController` 加能力后,Vue 壳(components/ExcelViewer.vue)与 React 壳(react/ExcelViewer.tsx)都要接上,e2e 各自覆盖。
- **默认只读、零回归**:`editable` 关闭时行为与历史完全一致;编辑能力(单元格/样式/列宽行高/图片/增删行列/公式重算/导出回写)是 **opt-in**,全建在框架无关 core 的命令栈 + 前后快照事件上(见 README「编辑」章节)。
- **e2e 浏览器**:`@playwright/test` 固定 `1.58.0`(对应本机缓存 chromium-1208,避开需下载的新版)。

## 常用命令

```bash
npm run dev          # 本地 demo(localhost:5300)
npm test             # 单元测试(node)
npm run test:e2e     # 真浏览器 e2e(Playwright;先 npx playwright install chromium)
npm run typecheck    # vue-tsc --noEmit
npm run build        # 构建库(dist/ 三入口 core.js+index.js+react.js + style.css + .d.ts;不打包 vue/react/exceljs/echarts/jspdf)
npm run dev          # Vue demo: localhost:5300/  ; React demo: localhost:5300/react.html
node scripts/gen-sample.mjs   # 重新生成 public/sample.xlsx
```

## 路线图(多框架架构 + 分包 + 发布)

- **Phase A ✅** 抽框架无关 ViewerController(叠加层 / 渲染引擎 / 选区+交互 / 查找+筛选+排序 / 导出编排)
- **Phase B ✅** 分包:单包三子入口(core/vue/react 共享 dist/core.js)
- **Phase C ✅** React 薄壳(src/react,共用 core,带真浏览器 e2e)
- **Phase D ✅** 文档:ARCHITECTURE / CONTRIBUTING / README(React 用法 + 三入口 + props/导出/编辑表)
- **编辑能力 E0–E8 ✅(→ 1.0.0)** 配置/只读 → 写数据层+命令栈+前后快照事件 → editor 扩展 → 内置编辑器 → 数据层语义统一(resize 入栈+脏状态/还原)→ 公式重算(可换引擎)→ 样式 → 图片 → 增删行列 → 导出 xlsx/json/csv
- **保真增强 F1–F3 ✅(→ 1.1.0)** 增删行列公式引用重写 / 图片导出 twoCell+EMU 偏移 / xlsx overlay 高保真
- **Phase E**(进行中)发布:`ooxml-excel-editor` 1.1.0 → `npm publish`(2FA)
- 仍未做:真正 workspace 三包拆分(目前单包多入口已够用);透视表 UI;大文件编辑性能

每阶段测试 green + 提交,不破坏现有 Vue / React、不破坏「默认只读零回归」。
