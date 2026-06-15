# ooxml-excel-editor — 项目开发准则(AI 与贡献者必读)

> Vue3 + Vue2 + React 高保真 .xlsx 预览/编辑组件。从零实现解析与 Canvas 渲染。**默认只读、零回归;开 `editable` 进入编辑**(值/样式/列宽行高/图片/增删行列/公式重算/导出回写 .xlsx·JSON·CSV)。

## ★ 七条中心原则(任何后续开发都要围绕,不得破坏)

1. **好文档** —— 同时服务**调用方**(怎么用 → README)和**二开者**(怎么改/扩展 → EXTENDING.md 扩展点 API + ARCHITECTURE.md 内部结构)。改公开 API/扩展点/导出,必须同步更新 README + 相关文档(EXTENDING/ARCHITECTURE/CONTRIBUTING/各包 README + CHANGELOG)。README 的 props 表 / 导出表 / 选项表要与代码一致。**四入口(`ooxml-excel-editor` / `/react` / `/vue2` / `/core`)统一 `export * from core` 同源** —— 加 core 公共出口自动各入口可见,别再各入口手维护清单(会漂移,1.14.1 因此修过)。
2. **可发布** —— 始终保持能 `npm publish` 的状态:exports/main/module/types 对、`.d.ts` 完整、peer 依赖(vue/react/exceljs 必需,echarts/jspdf 可选)**绝不打包进产物**、占位元数据要清。
3. **三壳共存** —— `core` 框架无关,Vue3/Vue2/React 各是薄壳共享同一 `core`。**禁止在 `core` 里 import vue/react**。新功能优先做进 `core`(框架无关),壳只做桥接。
4. **包名清晰** —— 现状: **单包四子入口**(`ooxml-excel-editor` = Vue 3 壳 / `/react` = React 壳 / `/vue2` = Vue 2 壳 / `/core` = 框架无关引擎)。Vue 3 / React / core 共享同一 `dist/core.js`;Vue 2 因 SFC 编译器冲突独立打包(内嵌 core)。后续生态大了可平滑拆成真正的 workspace 多包,现阶段单包多入口已满足"按框架各取所需",不过度拆包。
5. **扩展点** —— 保留并尊重已有扩展点:`:theme` / `transformModel` / `cellStyle` / 事件 / `overlay` slot / `rectOf` 命令式 API。新增能力优先做成可配置/可覆盖,而非写死。跨框架时把框架特定扩展点(如 Vue 3 `overlay` slot 返回 VNode)做成框架无关(返回 DOM/描述)。
6. **插件机制** —— `definePlugin` 打包 theme/transformModel/cellStyle/events/overlay/toolbar/setup;多插件按数组合并、组件 props 最后覆盖。改动不得破坏此契约;插件应跨框架可用。
7. **UI 1:1 复刻 ★(2026-06-08 新增)** —— **Vue 3 SFC 是参考实现 (Standard)**,任何 UI 变动 (props / events / 工具栏 / 公式栏 / 状态栏 / dialog / 浮层 / 演示 demo) 都**先在 Vue 3 上落地**,再**1:1 复刻**到 Vue 2 + React 壳。三个壳 + 三个 demo 必须视觉/交互一致 (类名 / 颜色 / 字号 / padding / 按钮顺序 / 下拉行为 全部对齐)。共享纯 TS/CSS 资源放 `src/demo-shared/` 或 `src/components/toolbar-icons.ts` 等中性目录,Vue 2 / React 直接 import。改 Vue 3 而漏改 Vue 2 / React 视为破坏此原则的 bug。

## 架构(framework-agnostic core + 三薄壳)

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

src/components(Vue 3 壳 ★参考实现)→ dist/index.js:ExcelViewer.vue + 子 SFC
                                      (ViewerToolbar/ActionToolbar/ToolbarMenu/FormulaBar/
                                       SheetTabs/ExportDialog/ExportProgressOverlay/FindBar/...)
                                      薄壳, onMounted 起 ViewerController, hooks 桥接响应式
src/react(React 壳, 复刻 Vue 3)→ dist/react.js:ExcelViewer.tsx + ExportProgressOverlay.tsx
                                  + 复用 src/components/toolbar-icons.ts / toolbar-types.ts
                                  薄壳, useLayoutEffect 起同一 ViewerController, useReducer force 桥接
src/vue2(Vue 2.6 + 2.7+ 壳, 复刻 Vue 3)→ dist/vue2.js:ExcelViewer.ts (render function, 不走 SFC)
                                + 复用 src/components/toolbar-icons.ts
                                独立打包(内嵌 core, 因 Vue 2/3 SFC 编译器冲突)
                                **Composition API 从 '@vue/composition-api' import** (同时兼容 2.6 plugin / 2.7 内置)
                                canvas/overlays/scroller 用 createElement + appendChild (imperative DOM)
                                避免 Vue 2 patch 重建 controller 持有的 DOM (vnode key 必须全设)
src/demo-shared(三 demo 共享 CSS / 工具)→ demo-bar.css (绿色头) + demo-editor.ts
```
数据流:文件 → loader → parser → **中间模型(WorkbookModel)** → ViewerController(含 CanvasRenderer)→ canvas。
中间模型与 ExcelJS/XML 形状**完全解耦**。`ViewerController` 是壳与引擎的唯一桥:壳给 DOM 元素 + 一组回调 hooks
(onRenderer/onRenderTick/onSelectionChange/onCellClick/…),控制器回调驱动壳的响应式重渲。**三壳逻辑同构、共用 ~100% 引擎**。

## 不可破坏的硬约束

- **测试是回归网**:改动后 `npm run typecheck` + `npm test`(单测)+ `npm run test:e2e`(Playwright 真浏览器)+ `npm run build` 必须全绿。当前基线 **419 单测 + 192 e2e(Vue 3 / React / Vue 2 三壳覆盖;Vue 2 e2e 跑独立 5302 dev server,见 `e2e/vue2-smoke.e2e.ts`)**。
- **core 不依赖框架**:`src/core/**` 不得出现 `from 'vue'` / `'react'`(构建后 `dist/core.js` 也不得 import vue/react/hyperformula/exceljs —— 重依赖全动态懒加载)。
- **三壳同构**:给 `ViewerController` 加能力后,Vue 3 壳(components/ExcelViewer.vue)、React 壳(react/ExcelViewer.tsx)、Vue 2 壳(vue2/ExcelViewer.ts)都要接上,各自 e2e 覆盖。**任何 UI 变更先 Vue 3 落地, 再 1:1 复刻到 Vue 2 + React**(详见第 7 中心原则)。
- **默认只读、零回归**:`editable` 关闭时行为与历史完全一致;编辑能力(单元格/样式/列宽行高/图片/增删行列/公式重算/导出回写)是 **opt-in**,全建在框架无关 core 的命令栈 + 前后快照事件上(见 README「编辑」章节)。
- **e2e 浏览器**:`@playwright/test` 固定 `1.58.0`(对应本机缓存 chromium-1208,避开需下载的新版)。
- **Vue 2 兼容底线 = Vue 2.6.12 ★(2026-06-08 钉死)** — `peerDependencies.vue` 声明 `^2.6.0 || ^2.7.0 || ^3.4.0`, 实际消费方就有 Vue 2.6.12 项目. 任何写进 `src/vue2/**` 的代码、任何文档里给 Vue 2 用户的示例, **不准用 Vue 2.7+ 才有的特性**, 包括但不限于:
    - ❌ **函数 ref / callback ref** (`h('div', { ref: (el) => ... })`) — Vue 2.7 backport, **Vue 2.6 完全忽略**. 必须用 string ref + `vm.$refs[name]` (见 [src/vue2/ExcelViewer.ts](src/vue2/ExcelViewer.ts) 的 `makeDomSlotFactory`).
    - ❌ **Vue 2.7 内置 Composition API** 直接 `import { ref } from 'vue'` — 必须 `from '@vue/composition-api'` (它在 2.6 走 plugin, 在 2.7 自动 re-export).
    - ❌ `<script setup>` (Vue 3 only, 跟 2.x SFC 不兼容).
    - ❌ Vue 2.7+ 才有的 `defineComponent` 类型推断细节假设.
    - ✅ 用 string ref / `Vue.use(VueCompositionAPI)` plugin / `defineComponent` (Composition API 入口) / `h()` render function. 这些三个版本 (2.6 / 2.7 / 3) 都受支持.
    - 验证手段: 把 `vue2` devDep 临时改成 `npm:vue@2.6.12` + 跑 `npm run dev:vue2` 复现 (`vue2-demo/main.ts` 需 `Vue.use(VueCompositionAPI)`, vite alias 也得改). 详见 CHANGELOG 1.3.2 二次迭代条目.
    - **背景**: 1.3.2 第一次发布前消费方 Vue 2.6.12 项目报 "renderArea DOM 拿不到", 根因正是函数 ref 在 2.6 被忽略. 当时 `vue2-demo` 跑 Vue 2.7.16 没暴露. 此后**新功能必须在 Vue 2.6 上验证**, 不要假设 Vue 2.7+ 的 API 可用.

## 常用命令

```bash
npm run dev          # Vue 3 demo (port 5300, 默认)
npm run dev:vue3     # Vue 3 demo (alias)
npm run dev:react    # React demo (port 5301)
npm run dev:vue2     # Vue 2 demo (port 5302, root=vue2-demo/)
npm test             # 单元测试(node, 419 个)
npm run test:e2e     # 真浏览器 e2e(Playwright;先 npx playwright install chromium)
npm run typecheck    # vue-tsc --noEmit
npm run build        # 构建库(dist/ 四入口 core.js+index.js+react.js+vue2.js + style.css/vue2.css + .d.ts;不打包 vue/react/exceljs/echarts/jspdf)
node scripts/gen-sample.mjs   # 重新生成 public/sample.xlsx
```

## UI 1:1 复刻工作流(第 7 中心原则的具体落地)

任何 UI 类改动按以下顺序进行,**漏跳任一步算 bug**:

1. **Vue 3 SFC 先落地** (`src/components/ExcelViewer.vue` 或子 SFC):写 template + scoped CSS + setup 逻辑。这是参考实现 (Standard)。
2. **抽共享资源**(如果有可跨框架复用的部分):
   - 纯 TS 模板/图标/类型 → `src/components/toolbar-icons.ts` / `toolbar-types.ts` / `export-types.ts`
   - 全局 CSS(三 demo 共享绿色头等) → `src/demo-shared/*.css`
   - 这些文件**禁止 `import 'vue'` / `'react'`**(中性 TS/CSS)
3. **复刻到 React 壳** (`src/react/ExcelViewer.tsx`):
   - JSX 结构 + className 用同款类名(或自己加 `rxl-` 前缀但保持视觉一致)
   - useState / useEffect 桥接同款行为
   - 复用步骤 2 抽出的资源
4. **复刻到 Vue 2 壳** (`src/vue2/ExcelViewer.ts`):
   - render function (h() + VNode) 写同结构;**每个 VNode 必须有 `key`**(否则 Vue 2 patch 会复用 DOM,见 ba6470c)
   - canvas / overlays / scroller / editor-slot 等 controller 持有的 DOM **必须 onMounted createElement + appendChild**(不交给 Vue patch)
   - 状态/computed/watch 用 Vue 2.7 Composition API
   - 复用步骤 2 抽出的资源
5. **三 demo 也要 1:1**(`src/App.vue` / `src/react-demo/main.tsx` / `vue2-demo/main.ts`):
   - 顶栏 demo 按钮列表对齐
   - 按钮顺序、文本、title 一致
   - 共享 demo-bar.css 类名(`.app-bar` / `.sample-btn` / `.file-btn` / `.edit-toggle`)
6. **验证**:`typecheck` + `npm test` + Playwright 截图三个 demo 视觉对比 + `npm run build` 全绿。

**已知 Vue 2 壳特殊坑**(在 [src/vue2/ExcelViewer.ts](src/vue2/ExcelViewer.ts) 头部 doc 详述):
- 函数 ref 每次 patch 会 invoke null + el (即使是同 DOM), controller 用 `instance` 上挂的非 reactive 字段存 DOM 引用
- Vue 2 把任意 object 转 reactive (包含 ResizeObserver), 用 `_xxx` 挂 instance 避开
- `updated()` 钩子内改 reactive 会触发 updated 死循环,用 `watch` 监听具体依赖代替
- `<template v-for>` Vue 2 需要 key 放在 iteree 上,不是 `<template>` 本身

## 路线图(多框架架构 + 分包 + 发布)

- **Phase A ✅** 抽框架无关 ViewerController(叠加层 / 渲染引擎 / 选区+交互 / 查找+筛选+排序 / 导出编排)
- **Phase B ✅** 分包:单包三子入口(core/vue/react 共享 dist/core.js)
- **Phase C ✅** React 薄壳(src/react,共用 core,带真浏览器 e2e)
- **Phase D ✅** 文档:ARCHITECTURE / CONTRIBUTING / README(React 用法 + 三入口 + props/导出/编辑表)
- **编辑能力 E0–E8 ✅(→ 1.0.0)** 配置/只读 → 写数据层+命令栈+前后快照事件 → editor 扩展 → 内置编辑器 → 数据层语义统一(resize 入栈+脏状态/还原)→ 公式重算(可换引擎)→ 样式 → 图片 → 增删行列 → 导出 xlsx/json/csv
- **保真增强 F1–F3 ✅(→ 1.1.0)** 增删行列公式引用重写 / 图片导出 twoCell+EMU 偏移 / xlsx overlay 高保真
- **WPS 单元格内嵌图 DISPIMG ✅(1.2.0)** 三期完成 + UX 打磨:① 解析 `xl/cellimages.xml` + `=DISPIMG()` 公式 → 画进格内展示,贴合方式可配置(`cellImageFit` fill/contain/cover,默认 fill 同 WPS);② 编辑模式互转 —— 就近嵌入(`convertImageToCellAuto`/几何反推)、整表/整列批量(`convertAllImagesToCells`,单次撤销)、嵌入→浮动,右键菜单接入,全入命令栈;③ **导出回注** —— ExcelJS 写出后在 zip 层回注 cellimages.xml + rels + media + Content_Types/workbook-rels 补丁(`export/wps-cellimages.ts`),rebuild/overlay 两模式都让导出 .xlsx 往返 DISPIMG。生成测试件:`node scripts/gen-wps-sample.mjs` → `public/wps-dispimg-sample.xlsx`。
- **WPS 风格长文本编辑 ✅(1.2.1)** 默认编辑器从 `<input>` 改 `<textarea>` + `wrapLines()` 算 desired height + onResize 钩子让 editor-host 浮起撑高;vAlign overflow fallback 修"提交后定位最下面"问题;公式栏 textarea + auto-resize 跟同款。
- **Vue 2 壳 + 三壳 UI 1:1 ✅(1.3.0)** ① Vue 2.7 壳(`/vue2` 入口)render function 实现, 共用 core;② 修了 Vue 2 patch 复用 DOM 致 controller stale 的根因(key + imperative DOM);③ Vue 3 / Vue 2 / React 三壳 UI 1:1 复刻 ViewerToolbar + ActionToolbar (SVG 图标 + 9 项内置 + 下拉子菜单) + StatusBar + Tooltip + ExportDialog + ExportProgressOverlay + overlay scoped slot;④ 三 demo 入口绿色头共享 `src/demo-shared/demo-bar.css` 视觉对齐 + 完整演示按钮 (JSON 示例 / PDF 水印 / 数据→JSON / EditTargets dialog / 高亮只读 / 编辑模式按钮组 / ⋯ 更多 溢出折叠);⑤ 独立 dev scripts (`dev:vue3` 5300 / `dev:react` 5301 / `dev:vue2` 5302). **★ 自此 UI 1:1 复刻成为第 7 中心原则**。
- **Phase E**(进行中)发布:`ooxml-excel-editor` 1.3.0 → `npm publish`(2FA)
- **透视表完整闭环 ✅(1.4.0)** **整个功能由 `pivotTable` prop 开启(三壳同名,默认 false 关闭 = 零回归;三 demo 已开启)**。① WPS 式创建入口:工具栏 `pivot-table`(开关关闭时不渲染)→ 选区 → 生成位置对话框(现有表单元格 / 新建表)→ 静态透视汇总表入命令栈;② 右侧「数据透视表」字段面板(core 框架无关 DOM,`viewer/pivot-dialog-host.ts`):搜索 + 按钮/拖拽进 筛选器/列/行/值 四区 + 筛选值下拉 + 汇总方式切换,每次变更重建结果;③ 编程 API `createPivotTable`/`createPivotTableFromSelection`/`openPivotTableDialog`,三壳 + 插件 viewer 暴露(开关关闭时返 false + 提示);④ **导出回注真实 OOXML pivot 零件**(`export/pivot-tables.ts`,同 cellimages 模式):pivotCacheDefinition/Records + pivotTableDefinition + workbook pivotCaches + 全套 rels + Content_Types,`refreshOnLoad=1` 让 Excel/WPS 打开即识别真透视表;筛选语义对齐 WPS("=值"写 pageField@item 还原选中、"非空"=多选+隐藏空白项);**overlay 导出从原件 zip 原样搬运原有透视表零件**(`restoreOriginalPivotPartsIntoZip`,按表名重挂 worksheet 关系,编号/cacheId 与新建的自动避让);⑤ pivot-parser 支持标准 rels 隐式关联(真 Excel 文件 + 导出件往返);⑥ **活刷新**(源数据编辑/撤销后透视表按 source 自动重算,唯一入口 recomputePivot,派生态不入命令栈)+ **行分组折叠/展开**(≥2 行字段产出大纲,canvas [−]/[+] 按钮 + pivotToggleAt 命中)+ **多选筛选**(PivotFilterMode include,面板勾选 + 导出 multipleItemSelectionAllowed/item@h)。
- **数据验证完整化 + Vue 2 e2e 回归网 ✅(1.8.0)** ① 数据验证从"只读取/list 选值"做到**编辑拦截**:解析全类型规则(list/whole/decimal/date/time/textLength/custom + operator + 出错/输入提示)进 `SheetModel.dataValidationRules`,框架无关引擎 `edit/data-validation.ts` 在提交时校验(stop 硬拒+模态、warning/info 软提示、空值/公式放行),框架无关 `validation-prompt-host` 出错模态/toast + 输入提示气泡(三壳共用);顺手修内置编辑器拒绝后 `done` 锁卡死的 UX bug(`commit()` 返成功与否)。② Vue 2 壳补 e2e(此前零覆盖):`playwright.config.ts` 加 5302 第二 dev server,`vue2-demo` 挂 `window.__excelViewerVue2`,`e2e/vue2-smoke.e2e.ts` + data-validation 加 Vue 2 行。
- **条件格式可编辑 ✅(1.9.0)** 整个功能由 `conditionalFormat` prop 开启(三壳同名,默认 false = 只读渲染、零回归;三 demo 已开启)。① 模型 `ConditionalRule` 加 id/origin/dirty/raw + top10/iconSet.reverse,`parseConditional` 派 id+存 raw;② 命令 `set-conditional`(整张数组替换,整体单次撤销);③ 编程 API getConditionalRules/add/update/remove/setConditionalRules/openConditionalFormatDialog(控制器+插件+三壳);④ 框架无关管理对话框 `viewer/conditional-format-dialog-host.ts`(三壳共用,6 类编辑器:cellIs/expression/colorScale/dataBar/iconSet/top10)+ 工具栏入口 `conditional-format`;⑤ 导出回写 `xlsx-writer` 的 writeConditionalFormatting:**未编辑的 parsed 规则用 raw 原样回写(零退化)**,用户新建/编辑的按模型 buildExcelCfRule 构造;rebuild + overlay 都回写(rebuild 此前丢弃)。
- **自动填充柄 ✅(1.10.0)** Excel/WPS 拖拽填充。纯框架无关 core canvas 交互(`edit/autofill.ts` 序列引擎 + `canvas-renderer` 画柄/命中/虚线预览 + 控制器 `fill` 拖拽模式 `setCellsBatch` 单次撤销),三壳零改动自动获得;需 `editable`。序列:数值等差/日期/前缀+末尾整数文本/星期月份循环/兜底复制 + Ctrl 翻转复制↔序列。v1 填值不复制格式。
- **查找替换 + 数字格式编辑器 + 批注编辑 ✅(1.11.0)** 三个编辑小件合并。① 查找替换:控制器 setFindReplace/replaceCurrent/replaceAll(全部替换单次撤销),三壳查找栏加替换行(editable 才显示);② 数字格式编辑器:框架无关 `viewer/number-format-dialog-host.ts`(分类 + 预览复用 number-format 引擎 + 自定义代码)+ 工具栏 `number-format` 入口 + setSelectionNumberFormat;③ 批注编辑:`set-comment` 命令 + `model/mutations.ts` setCellComment + 框架无关 `viewer/comment-dialog-host.ts` + 右键菜单 + 导出回写 ExcelJS note(rebuild + overlay)。
- **格式刷 ✅(1.12.0)** Format Painter。纯框架无关 core 交互:控制器 `startFormatPainter`(采样活动格完整样式)/ `isFormatPainterArmed` / `cancelFormatPainter`,刷动作在 `onMouseUp` 选区完成后 setStyle(单次撤销);工具栏 `format-painter` 入口(active 态反映待刷)+ copy 光标 + Esc 退出。三壳 + 插件 ViewerApi 暴露。
- **Ctrl 多区域选择 ✅(1.13.0)** 不连续多选:选区模型加 selRanges[] + getSelectionRanges/hasMultiSelection,onMouseDown Ctrl 加选/非 Ctrl 清,renderer setExtraSelection 画所有区(多选不画填充柄),copyMultiSelection 逐行堆叠 TSV+HTML,getSelectionStats 跨区聚合(三壳状态栏改用)。纯 core,三壳零改动(壳只转发鼠标 + 状态栏改调 getSelectionStats)。
- **内置 MIT 公式引擎 ✅(1.14.0)** 从零实现 formula/builtin(parse 词法+优先级解析 → AST;eval 求值器+错误传播;functions ~60 函数;index 依赖图+拓扑级联+循环检测)。设为 recalc 默认引擎(替代 GPL HyperFormula,零依赖);HyperFormula 仍可经 :formula-engine 注入(hyperFormulaEngineFactory)。30 单测;现有 recalc e2e 改由内置引擎驱动仍过。同版**公式自动补全**(edit/formula-autocomplete.ts,框架无关默认编辑器内置,三壳自动有):输 =SU 弹函数名 + 参数提示(FUNCTION_SIGNATURES),Enter/Tab/点选 插入 NAME(。
- 仍未做(用户已挑但暂缓):大文件编辑性能(需用户给真实大文件 profiling)。其它:真正 workspace 多包拆分;Vue 2 子入口体积优化 (现 423 KB);rebuild 导出不搬运原文件只读透视表(仅 overlay)。

每阶段测试 green + 提交,不破坏现有三壳、不破坏「默认只读零回归」、**不破坏 UI 1:1 复刻**。
