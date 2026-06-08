# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与 [语义化版本](https://semver.org/lang/zh-CN/)。

## [1.3.2] - 2026-06-08

**修 webpack 4 / vue-cli 4 / CJS 环境兼容** — 用户反馈 Vue 2.6.12 + @vue/cli 4 (webpack 4) 项目消费 1.3.1 时报多个错. 此版本系统修复 4 个老打包器兼容问题, **任何环境(Vite / webpack 5 / webpack 4 / Snowpack / Parcel)都能消费**.

### 修复

#### 1. dist 不再含 `import.meta.url` + module worker (webpack 4 SyntaxError 致命)

之前 lib-vue2 build 没用 worker stub, `dist/vue2.js` 含:
```js
new Worker(new URL("/assets/parse.worker-xxx.js", import.meta.url), { type: "module" })
```
- `import.meta.url` 是 ESM 专属语法, webpack 4 解析失败 (Module parse failed: Unexpected token)
- `/assets/parse.worker-xxx.js` 是 Vite 构建产物, 在别的打包器/项目里根本不存在

**修复**: vite.config.ts 把 `isLibBuild` 判定从 `command === 'build' && !isDemoSite && !isVue2Build` 改成 `command === 'build' && !isDemoSite` (Vue 2 lib build 也用 worker-client.stub.ts 走主线程解析). 现在所有 lib 产物 (`dist/index.js` / `react.js` / `vue2.js`) 都不含 `new Worker` 也不含 `import.meta`.

#### 2. dist 降级 `??` / `?.` 等 ES2020 语法 (webpack 4 SyntaxError 致命)

之前用户报:
```js
r.onerror = () => o(r.error ?? new Error("文件读取失败"))  // ?? 是 ES2020, webpack 4 不识别
```
- 多数项目默认不转译 node_modules (webpack 4 / vue-cli 4 尤甚), 用户被迫加 transpileDependencies

**修复**: vite.config.ts 给所有 lib build 加 `target: 'es2018'`. ES2018 没有 `??` (ES2020) / `?.` (ES2020), 但保留 async/await + spread/rest (ES2017+, webpack 4 都支持), 不引入大量 polyfill.

#### 3. 包根加 vue2.js / react.js / core.js stub (webpack 4 不读 package.json#exports)

之前用户报:
```js
import ExcelViewer from 'ooxml-excel-editor/vue2'  // webpack 4 解析失败
// 被迫: import ExcelViewer from 'ooxml-excel-editor/dist/vue2.js'
```

**修复**: 包根新增 3 个 stub 文件:
- `vue2.js` — re-export `./dist/vue2.js`
- `react.js` — re-export `./dist/react.js`
- `core.js` — re-export `./dist/core.js`

webpack 4 解析 `ooxml-excel-editor/vue2` → 找包根 `vue2.js` → re-export 到 dist. webpack 5 / Vite / Rollup 仍走 `package.json#exports` → `./dist/vue2.js` (stub 不会被用到, 0 性能影响).

文件加进 `package.json` `files` 数组, npm publish 时一并发布.

#### 4. echarts / exceljs / jspdf / hyperformula **inline 进 dist chunks/** (彻底)

试过 3 个方案都不彻底:
- 1.3.1 `peerOptional`: npm 7+ 触发 ERESOLVE 冲突
- 1.3.2 第一版 文档说明 + 用户按需装: webpack 4 静态扫描 `await import()` 找不到模块发 warning
- 1.3.2 第二版 改 `dependencies` 自动装: webpack 4 解析这些 lib **源码** 时仍报 named-export / class-fields 错误 (因为 lib 用 ES2020+ 语法)

**终极修复**: rollupOptions.external 区分两类:

**inline 编进 chunks/** (老打包器零解析):
- `exceljs` — 私有依赖, 用户项目一般没用
- `hyperformula` — class fields 语法 webpack 4 解析必炸, 必 inline
- `jspdf` — 同样现代语法

**external + dependencies (npm 自动装)**:
- `echarts` — 用户项目大概率已有自己的 echarts (常见 viz lib), inline 会:
  - 浪费 ~1 MB bundle 体积
  - 跟用户 `echarts.registerTheme()` 注册的主题 dual instance, 用户主题在我们 inline 的实例上失效
  现代 echarts 5.x CJS / UMD 入口 webpack 4 兼容良好, external 安全.

target: `es2017` 把 class-fields / `??` / `?.` 全降级.

消费方:
- **完全不用手动装任何 lib** — `npm i ooxml-excel-editor` 自动装 echarts
- webpack 不解析 exceljs / jspdf / hyperformula 源码 (已 inline 嚼碎)
- echarts 跟消费方自己的 echarts 同实例 (theme 共享, 不 dual)
- 不再有 named-export / class-fields / `import.meta` / module worker / `??` 等老打包器报错

变化:
- **dependencies** 仅保留 `echarts` (+ 小 lib `fast-xml-parser` / `fflate`)
- **dependencies** 移除: `exceljs` / `jspdf` / `hyperformula` (已 inline)
- **peerDependencies** 剩余: `vue` / `react` / `react-dom` / `@vue/composition-api`
- **dist 体积**: 5.4 MB (含 3 个 inline lib)
- **tgz 压缩后**: 1.33 MB
- vite/rollup target: `es2017`

消费方 build 后 bundle (不用 PDF/编辑公式) **影响很小** — dynamic import 保留 code-split, 不调到 chunk 就不下载.

### 总效果

任何打包器都能消费 1.3.2:
- ✅ Vite / webpack 5: 走 `package.json#exports`, 用 dist/ 真实产物
- ✅ webpack 4 / vue-cli 4: 走包根 stub 文件, dist 已降级语法 + 无 worker
- ✅ Parcel / Snowpack / esbuild: 同上
- ✅ Node CJS (SSR): 同 webpack 4 路径

### 升级提示 (从 1.3.0 / 1.3.1 升 1.3.2)

无 breaking, `npm i ooxml-excel-editor@1.3.2`. 4 个运行时 lib (`echarts` / `exceljs` / `jspdf` / `hyperformula`) 已改成 `dependencies` 随包自动装, **不用手动 `npm i`**. 之前手动装的不用卸载, npm 会复用现有版本 (前提是 `^range` 满足).

Vue 3 / React 用户**无变化**.

---

## [1.3.1] - 2026-06-08

**Vue 2 子入口扩展兼容到 Vue 2.6.x** — 通过 `@vue/composition-api` package 让一份代码同时支持 Vue 2.6 + 2.7+. 1.3.0 只支持 2.7+ (内置 Composition API), 此版本不破坏 2.7 用户用法.

### 新增

- **Vue 2.6.x 支持**: peerDeps `vue` 从 `^2.7.0 || ^3.4.0` 扩到 `^2.6.0 || ^2.7.0 || ^3.4.0`
- **`@vue/composition-api` 加入 peerDependencies** (optional peer): Vue 2.6 用户必装 + `Vue.use(VueCompositionAPI)`;Vue 2.7+ 用户也装 (它自检测内置 API,plugin 是 noop,无需 Vue.use)

### 实现

- Vue 2 壳代码 (`src/vue2/ExcelViewer.ts` + `use-excel-document.ts`) 的 Composition API import 从 `'vue2'` 改成 `'@vue/composition-api'` (一份代码同时支持 2.6 plugin / 2.7 内置)
- `vite.config.ts` dev alias `@vue/composition-api` → `vue@2.7 dist`(dev 时拿 2.7 内置), build external (消费者自己解析)
- docs/Vue2.md 加 Vue 2.6 vs 2.7 安装步骤对比 + 解释为什么要装 `@vue/composition-api`

### 升级提示 (从 1.3.0 升 1.3.1)

**Vue 2 用户**: 必须额外装 `@vue/composition-api`. 1.3.0 装的用户升级时 `npm install` 会提示缺失 peer, 跟着装即可:

```bash
npm i @vue/composition-api
# Vue 2.6.x 还需 main.js: Vue.use(require('@vue/composition-api').default)
# Vue 2.7+ 不需要 Vue.use (plugin 自动 noop)
```

Vue 3 / React 用户**无变化**.

---

## [1.3.0] - 2026-06-08

**Vue 2 兼容子入口 + 三壳 UI 1:1 复刻 + 独立 dev scripts** — Vue 2.6 / 2.7+ / Vue 3 / React 三个壳视觉与交互完全一致 (Vue 3 SFC 是参考实现 Standard, Vue 2 / React 1:1 复刻).

### 新增

- **`ooxml-excel-editor/vue2` 子入口** (`src/vue2/ExcelViewer.ts`, ~1000 行 render function 版): **Vue 2.6.x + 2.7+** 兼容壳, 跟 Vue 3 壳 1:1 功能对齐. 通过 `@vue/composition-api` package 同时支持 2.6 (装 plugin + `Vue.use`) 和 2.7+ (plugin 自检测 + noop)
  - 全部 28 项 props (src / workbook / jsonOptions / templateFile / fileName / theme / cellStyle / cellImageFit / imageLightbox / openLinks / plugins / toolbar / editable / cellReadOnly / readOnlyRanges / editableTargets / strictDimensions / readOnlyCellStyle / editor / recalc / formulaEngine / contextMenu / exportProgress / transformModel / templateName)
  - 全部 15+ events 跟 Vue 3 同名 (rendered / error / progress / cell-click / cell-dblclick / selection-change / sheet-change / hyperlink-click / cell-change / edit-start / edit-commit / dim-change / dirty-change / image-change / struct-change / permission-denied / before-context-menu / context-menu)
  - 完整命令式 API (80+ 方法) 跟 Vue 3 `viewerApi` / React `ExcelViewerHandle` 对齐
  - 插件系统完整 (events / cellStyle / theme / overlay / toolbar / setup / contextMenu / editor)
- **独立 dev scripts**: `npm run dev:vue3` (port 5300, 默认) / `npm run dev:react` (port 5301) / `npm run dev:vue2` (port 5302), 三个 demo 进程隔离 + 独立端口

### UI 1:1 复刻 (★ 新中心原则)

**Vue 3 SFC 是参考实现 (Standard)**, 任何 UI 变动先 Vue 3 落地再复刻到 Vue 2 / React. 三壳 + 三 demo 视觉/交互对齐:

- **顶部 ViewerToolbar**: 文件名 + " · 模板: <name>" + 灰色 "N 个工作表" + "导出 ▾" 下拉 (PNG / PDF 位图 / PDF 矢量 / 打印 / 导出设置…) + 缩放 [− / select / +]
- **Action ToolBar**: 完整 9 项内置工具 (find / filter / clear-filter / copy / wrap-text / template / image-tools / freeze / export / zoom) + SVG 图标 (`src/components/toolbar-icons.ts`) + 下拉子菜单 + separator 分组
- **状态栏**: 选区范围 (A1:E5) + 计数 / 求和 / 平均 / 最大 / 最小 (调 `renderer.selectionStats`)
- **cell tooltip**: `onTooltip` hook → ref<TooltipState>, 支持 default / comment 黄底批注样式
- **ExportProgressOverlay**: 居中模态 + stage 标签 + 进度条 + 取消按钮 + `chain()` 包装所有 export API + `:export-progress` prop
- **ExportDialog**: 高级导出配置 (范围 / 清晰度 / 内容 / PDF 类型 / 纸张) — 工具栏 "导出设置…" 弹出
- **overlay scoped slot**: `<template v-slot:overlay="{ rectOf, rectOfRange, tick }">` 跨壳同 API
- **loading / error / empty 三态浮层**: progress bar + 错误提示 + 空状态
- **三 demo 入口** (App.vue / vue2-demo / react-demo) **绿色头共享** `src/demo-shared/demo-bar.css`:
  - 共同按钮: 选择 .xlsx / 加载示例 / JSON 示例 / 编辑模式 / PDF(页码+水印) / 数据→JSON / 贴合 select / ↓XLSX / ↓CSV / ↓JSON
  - 编辑模式按钮组: 设置可编辑 (EditTargets 白名单 dialog) / 高亮只读 / B 加粗选区 / 合并 / 拆分 / 背景 / 字体 / 清除填充 / 整表嵌入 / 格→图 / +行 / -行
  - "⋯ 更多" 溢出折叠 (ResizeObserver + 测量行)

### 修复

- **Vue 2 patch 复用 DOM 致 controller stale 引用 (致命空白渲染 bug)**: Vue 2 patch children 算法没 key 时按 tag 匹配, 把 `.ov-render-area` div 复用成 `.ov-toolbar` (改 className + 替换内容), controller.els.renderArea 变 stale (指向 toolbar 元素), 内部 canvas/overlays 都被 Vue 一并清掉. **修复**: 所有 chrome 子节点加唯一 `key` + canvas/overlays/scroller/editor-slot 改 imperative DOM (onMounted createElement + appendChild, Vue 完全不碰)
- **chrome computed 读 renderTick 导致滚动卡顿**: `fbCanEdit` / `formulaBarEditString` / `renderActionToolbar` 读了 `renderTick.value`, 每帧 scroll → root render function 重跑 → 整个 chrome DOM patch. **修复**: 只依赖低频 selVersion/findVersion/filterVersion
- **`updated()` 钩子内改 reactive → 死循环**: `demoRemeasure()` 改 `demoItemWidths` 触发 re-render → 又 `updated()` → CPU 100% 卡死. **修复**: 改用 `watch` 监听具体根本依赖
- **rAF 等浏览器 layout**: chrome 刚加上时 `renderArea.clientHeight` 是中间态 36px (不是最终 537px), `await nextTick()` 后加 `await rAF` 保证 layout 完成再 measure

### 已知限制

- `dist/vue2.js` 424 KB (含内嵌 core) — Vue 2 build pass 独立, 不共享 `chunks/`。后续探索 rollup 多入口共享 chunk
- `dist/vue2.d.ts` 类型声明暂未生成 (vue-tsc 不认 Vue 2 SFC; 本入口是 .ts 后续可补)
- Vue 2 e2e 覆盖待补 (Playwright 多 entry 配置)

### 文档

- 新 [docs/Vue2.md](docs/Vue2.md) — Vue 2 子入口完整使用文档 + 跟 Vue 3 差异表 + Vue 2 壳特殊坑(函数 ref / ResizeObserver / updated 死循环 / template v-for key)
- [README.md](README.md) 加 Vue 2 入口章节 + 四个子入口对照表 + 三壳 UI 1:1 说明
- [ARCHITECTURE.md](ARCHITECTURE.md) 加 demo-shared / 三壳对齐 / 第 7 中心原则段
- [CLAUDE.md](CLAUDE.md) 加第 7 中心原则 (UI 1:1 复刻) + UI 1:1 工作流 + Vue 2 壳特殊坑 + 路线图更新

### Commits (累计 9 个)

`ba6470c` 修空白渲染 + dev scripts → `0f9718a` statusbar/tooltip/ExportProgressOverlay → `d74c9f1` ExportDialog/overlay slot → `e735e8f` Vue 2 1:1 工具栏 → `b075e05` React 1:1 工具栏 → `5462f5e` demo 绿色头 + 演示按钮 → `ee199c3` EditTargets dialog 迁移 → `411e328` Vue 2 ⋯ 更多溢出折叠 → `a312222` 修死循环

---

## [1.2.1] - 2026-06-08

**WPS 风格长文本编辑** — 默认编辑器从 `<input>` 升级 `<textarea>`, 输入长文本自动换行 + 向下撑高, 跟 WPS / Excel 用户习惯一致。

### 修复 + 改进
- **★ 默认编辑器换 textarea + 动态撑高(WPS 风格, 2026-06-08)** — 用户反馈 / 截图: 编辑长文本时, 旧 `<input>` 单行水平溢出被裁切, 看不到全文; WPS 同场景**编辑框向下浮起 + 自动换行**显示完整内容。修复:
  - [default-editor.ts](src/core/edit/default-editor.ts) `<input>` → `<textarea>`, `white-space:pre-wrap` + 行高同步渲染层 `LINE_HEIGHT_FACTOR=1.18`. 短文本 (`rows=1`) 视觉跟之前 `<input>` 一致, 不破坏短输入体验. **Shift+Enter** 插入换行, 普通 **Enter** 提交 (跟 Excel/WPS 一致).
  - 新钩子 `CellEditorReturn.getDesiredHeight(widthPx) → number` + `CellEditorContext.reposition()` — 编辑器 input 事件触发 reposition, host 重测撑高. 用 [text.ts 的 `wrapLines`](src/core/render/text.ts) 复用渲染层换行算法, 保证视觉一致.
  - [editor-host.ts](src/core/edit/editor-host.ts) `position()` 新逻辑: 宽度仍 = 列宽 (跟 WPS 一致, **仅向下溢出**); 高度 = `max(单元格原高, getDesiredHeight(列宽))`, 上限 viewport 一半 (防一格输入 10000 字撑爆屏).
  - [`.editor-slot`](src/components/ExcelViewer.vue) + [`.rxl-editor-slot`](src/react/excel-viewer.css) CSS `overflow: hidden → visible`, 让编辑器可溢出原格. z-index:6 仍最上层, 不影响下方网格 / 冻结窗格 / 滚动交互.
  - **提交后行高保持原样** (跟 WPS 一致, 不持久化撑高): 短期视觉浮起仅在编辑期, 提交后单元格还原. 如需永久撑高, 用户单元格设 `wrapText=true` 走已有 autofit (历史行为不变).
- **★ 公式栏 (fx 内容条) 长文本自动撑高 (2026-06-08)** — 用户反馈: 公式栏内容过长被 `text-overflow:ellipsis` 截断, 看不全公式 / 长文本. 修: Vue + React 壳公式栏 `<input>` → `<textarea>` + `auto-resize` (内容变化时 height = scrollHeight). 跟单元格编辑器一样, 普通 Enter 提交、Shift+Enter 插换行、上限 ~6 行 (max-height: 108px) 超过内部滚动. CSS 同步 — formula-bar 高度自动撑, addr / fx 区垂直居中.
- **★ vAlign 溢出 fallback 顶对齐 (跟 WPS 一致, 2026-06-08)** — 用户反馈: 输入长文本提交后, 单元格"显示文末而不是开头". 根因: [canvas-renderer.ts:1090](src/core/render/canvas-renderer.ts#L1090) 默认 vAlign='bottom' 时, 文本总高超过单元格高 → `startY = y + h - pad - totalH` 变负, 首行画到格外, 用户看到的是最后几行. 修: 当 `totalH > availH` 时, 强制走 `'top'` 分支 (显示文头, 末尾裁切). 跟 WPS / Excel 行为一致.
- **`CellEditorFactory` 返回类型扩展** — 旧 `HTMLElement | { el, destroy? }` → 新增可选 `{ el, destroy?, getDesiredHeight?(w) }`. 旧自定义编辑器 100% 向后兼容 (不返 `getDesiredHeight` = 高度锁单元格高, 老行为).
- **`CellEditorContext.reposition()`** — 给自定义编辑器: 主动通知 host 重测撑高. 内容变化后调一次. 不调 = 不撑 (老行为).

### 测试
- 新单测 [default-editor.test.ts](src/core/edit/__tests__/default-editor.test.ts) 10 项: textarea 类型 / commit-cancel 行为 / Enter 移动 / Shift+Enter 换行 / 失焦提交 / 样式贴合 / `getDesiredHeight` 钩子存在 / 兜底 (依赖 canvas measureText 的具体撑高数测试走 e2e 覆盖)
- 新 e2e [edit-long-text.e2e.ts](e2e/edit-long-text.e2e.ts) Vue + React 双覆盖 12 项: 编辑器 textarea / 短文本不撑大 / 长文本撑高 / 动态变化 / Esc 取消 / Shift+Enter 换行 + Enter 提交
- devDependency 加 `jsdom` (单元格编辑器单测依赖 DOM 环境;**仅 dev**, 不影响 dist)

---

## [1.2.0] - 2026-06-08

**主线**: 只读边界三件套(白名单 / 尺寸多形态 / 视觉钩子 + permission-denied 事件)+
模板语义重设计(样式捐赠者)+ WPS 单元格内嵌图(展示 / 互转 / 导出往返)+ 富粘贴 +
图片放大下载 + 虚拟空行 + 公式栏可编辑 + 背景/字体色 + 编辑 UX 补齐(合并/粘贴/右键菜单)+
性能 + 导出错误可见性 + 1900/1904 日期单测。**全部向后兼容、默认只读零回归**。

测试基线: 306 单测 / 107 e2e。

### 修复 + 新增
- **★ 只读边界三件套(2026-06-08;三次提交)** —— 用户三连问"可编辑是一切改变的基础", 怀疑有路径绕过 isEditable. 三阶段解决:
  - **Phase A 闸门补漏 + `permission-denied` 事件**(commit 57ebdbc):审计找出 4 个绕过点 — `pasteRich` 合并粘贴 / `pasteRich` 图片粘贴 / `mergeCells/unmergeCells` 没逐格检查 / 图片转换族 没检查目标格.全部修复.行为 = 默认 skip(跟 `editRange` 一致)+ 一次操作结束 emit 一次 `permission-denied: { reason, cells[], dims?, message? }`(8 种 reason).Vue `@permission-denied` / React `onPermissionDenied`.右键菜单 disabled 完善:粘贴/清除/换行 = 选区有至少 1 格可编辑;合并/拆分/删除行列 = 全部可编辑.新 helper `partitionByEditable` / `rangeAllEditable` / `collectDeniedInRange`.
  - **Phase B 尺寸 API 多形态 + `strictDimensions`**(commit fac634d):用户问"现 API 是否能设置不相邻一批列宽行高".新类型 `DimTarget = number | number[] | { from, to }` 自动识别.升级 `setColumnWidth(target, width)` / `setRowHeight(target, height)` 接 union,**返回值 boolean → number 成功条数(BREAKING)**;新增 `autoFitColumns/Rows(target?)` / `resetColumnWidth/RowHeight(target)` 公开 API(target 不传 = 整表).多 index 时聚合成单次 `restore-wb` undo.新 prop `:strictDimensions` 默认 `false`:`true` 时该列/行至少 1 格在白名单内才能改尺寸,跟"白名单未覆盖 = 完全只读"严格语义一致.新 helper `canEditDimension` / `normalizeDimTarget`.
  - **Phase C 只读视觉钩子 + cursor**(本次):`CellStyleFn` 签名加可选第 3 入参 `ctx: CellStyleCtx { editable: boolean }`,旧 `(cell, pos) => ...` 完全兼容.新 prop `:readOnlyCellStyle: boolean | CellStyleOverride | CellStyleFn` — `true` 套内置浅灰 `#f5f7fa`,对象 = 固定样式,函数 = 按格自定义,默认 `false` 无视觉差异.canvas-renderer 渲染顺序:base style → cellStyle(传 ctx)→ 只读格再叠 readOnlyCellStyle.控制器注入 `isEditable` 回调给 renderer;鼠标悬停只读格(且编辑模式开)自动 `cursor: not-allowed`.Demo 加「高亮只读」toggle 跟「设置可编辑」配合使用.
  - **测试**:permissions.test.ts 共 27 项(11 项白名单 + 4 项 Phase A helpers + 7 项 Phase B canEditDim/normalizeDim);新单测 `readonly-style.test.ts` 7 项;3 个新 e2e spec — `permission-denied.e2e.ts` 4 项 / `dimension-targets.e2e.ts` 6 项 / `readonly-visual.e2e.ts` 3 项.总 281 单测 + 101 e2e 全绿.
- **★ 可编辑白名单 `editableTargets`(2026-06-08)** —— 现有的 `cellReadOnly` / `readOnlyRanges` 是**黑名单**(默认全可编辑,标只读),新加一个 **白名单**(默认全只读,标可编辑)。用户原话:"设置不相邻的一批单元格可编辑,暴露的 API 应该支持多种参数,可以传区域,也可以传单个单元格,多个单元格,行,列;没有设置的就是只读"。
  - **新类型** `EditableTarget`(`core/edit/types.ts`)联合 4 种形状,**自动识别**带哪些字段:`{row,col}` 单格 / `{row}` 整行 / `{col}` 整列 / `MergeRange` 矩形。单值或数组都接,**允许不相邻**多 target
  - **新 prop** `:editableTargets` —— `undefined` (不传) = 不启用白名单 = 老行为;`[]` (显式空) = 全只读;非空 = 白名单生效。优先级:`editable=false` ► 不在白名单 ► `readOnlyRanges` ► `cellReadOnly` ► 否则可编辑(白名单内仍能被黑名单二次"黑"掉)
  - **新命令式 API** `viewer.setEditableTargets(targets)` / `viewer.getEditableTargets()` —— 运行时改不动 prop,直接覆盖 `editCfg.editableTargets`,立即重绘
  - **demo 加「设置可编辑」按钮 + 弹窗**(编辑模式下出现):12×8 网格化点选,可单击格 / 列标题 / 行号切换;footer 显示已选数;「应用」生效 / 「关闭白名单」恢复默认 / 「取消」放弃
  - **测试**:`permissions.test.ts` 加 11 项白名单分支(单格 / 多格不相邻 / 整行 / 整列 / 矩形 / 空数组 / 混合 4 种 / 与黑名单叠加 / editable=false 时白名单也无效 / 形状识别);新 `e2e/editable-targets.e2e.ts` 3 用例(命令式 setEditableTargets 4 种 target 形状 / Demo 弹窗点选→应用 / 空数组全只读)
  - 不破坏现有 API,完全 additive
- **★ 模板语义彻底重设计(2026-06-08)—— 从"占位符 + 锚点表"改为"样式捐赠者"** —— 之前几轮"占位符 + 锚点表"语义反复改: `discardUnmatched` → `trimUnused` → JSON 自然位置… 都没切中用户实际需求。用户原话:**"模板这个功能,只在 json/csv 这些本身不附带格式的数据源才生效,xlsx 数据源根本不需要"**。新语义:
  - **`:templateFile` = 样式捐赠者** —— 模板贡献 styling(`styles` 池 / `merges` / 列宽 / 行高 / `freeze` / `theme`),JSON / CSV 数据**在 A1 自然位置渲染**,模板的 raw 文字 / 占位符 / 图 / 图表 / 条件格式 / 数据验证 **全部丢弃**(避免幽灵规则)
  - **仅在 `:workbook` (JSON / 模型) 数据源下生效** —— xlsx 数据源(`:src`)自带格式,给 `:templateFile` 会被**忽略并 console.warn**;工具栏「模板 ▾」在 xlsx 模式下**禁用**
  - **API 改动(BREAKING)**:删除 `props.template` (`TemplateFillSpec`)、删除 `viewer.applyTemplate(spec)` 命令式 API、删除 `core/template/fill.ts` (`fillTemplate` / `replacePlaceholders` / `TemplateAnchor` / `TemplateFillSpec`) 整个模块。新增 `core/template/style-overlay.ts` 的 `applyStyleTemplate(dataWb, templateWb)` 纯函数(同步,无 onProgress / signal)
  - **核心契约 9 项单测**(`style-overlay.test.ts`):数据 raw 全部保留在自然位置 / 模板 raw 全部丢弃 / styleId 从模板同位置取 / merges + 列宽 + 行高 + freeze 全部从模板拷贝 / 模板的 images + charts + conditional + dataValidations 不带过来 / 数据 sheet 名 + date1904 + cellImages 透传 / 入参不被原地修改 / 空数据 + 模板 → 干净样式骨架 / 数据 dimension 跟模板列宽行高声明取大
  - **demo 简化**:`App.vue` 删 `templateSpec` 计算属性(placeholders + anchors 整套不要),`loadJsonSample` 只传 `:workbook` 数据数组;`:fileName="'订单数据"`。点工具栏「模板 ▾ → 导入 .xlsx 模板」即套样式,模板的标题/客户/合计/{{占位符}} 全部不见, JSON 5 条数据仍在 A1 起的自然位置
  - **e2e 全重写**(`template-switch.e2e.ts`):3 个用例 —— ① JSON 无模板时数据在 A1 自然位置 ② JSON + 模板时 模板装饰文字全部不见 + JSON 仍在 A1 + 切换/清除模板正常 ③ xlsx 数据源加载后工具栏「模板 ▾」禁用
  - **过渡 / 后续迁移**:用户用旧 `applyTemplate(spec)` 自填数据的场景,改为"前端构造好完整 JSON 后再传 `:workbook`"。后续如需"模板填值"那套(不入命令栈的占位符 / 锚点表写入),应该走"应用层"而非组件 prop,组件只管渲染样式
- **右键菜单全面开放(Plan C)** —— 1.2.0 前右键菜单是内置 hardcoded、无任何对外接口。此版三层开放:
  - **prop `:contextMenu`** —— `false` 关闭内置弹层(事件仍触发,供自渲染);`(ctx, items) => MenuItem[] | undefined` transform 回调,在内置 items 上加 / 减 / 重排
  - **事件 `@before-context-menu` / `@context-menu`**(Vue) / `onBeforeContextMenu` / `onContextMenuShow`(React)—— `payload.preventDefault()` 接管渲染(用 Element Plus / Radix 等自渲染);`@context-menu` 拿到 `{x, y, ctx, items}` 总会触发
  - **命令式 API** `openContextMenu(x, y, items?)` / `closeContextMenu()` —— 键盘 Shift+F10、工具栏触发等
  - **插件贡献** `definePlugin({ contextMenu: (ctx, items) => ... })` —— 多插件按数组顺序串行,组件 prop 最后覆盖
  - 顺序固定:**内置 → 插件 → prop → 事件**;`MenuItem` / `ContextMenuCtx` / `ContextMenuTransform` 全部从 `/core` 导出
  - 顺手修了 Vue 4-pkt 接 prop 的"boolean prop 缺省被 Vue 判成 false"小坑(withDefaults 显式 undefined)
- **内置导出进度遮罩 + 三层覆盖机制(P1.5)** —— P1 已建好 `onProgress` + `AbortSignal` 协议,但壳没接 UI,用户调 `viewer.downloadPdf()` 看不到任何反馈 → 这次补 **Shell 默认 UI** + 用户**逐层覆盖**:
  - ① 新组件 `ExportProgressOverlay`(Vue SFC + React tsx)居中模态:stage 标签 + 进度条 + 取消按钮;同视觉 / 同协议
  - ② 壳自动 wrap 长任务(`downloadPdf`/`exportPdf`/`downloadImage`/`exportImage`/`downloadXlsx`/`exportXlsx`/`print`/`convertImagesInRangeToCell`/`convertCellImagesInRangeToFloat`):建内置 `AbortController` + 接 `onProgress` → 用户调时**默认看见遮罩**,无需任何 prop
  - ③ 用户传 `{ onProgress, signal }` 仍正常**链回调**(并存)
  - ④ 覆盖:`:export-progress="false"` 关闭内置遮罩(纯回调);Vue `#export-progress` 插槽 / React `renderExportProgress` 自渲染(拿到 `{state, busy, cancel}`)
  - ⑤ **修单表 PDF 卡 0% bug**:核心导出对"单表"和"jsPDF/canvasToBlob 黑盒"阶段改 emit `ratio: undefined`,overlay 走 indeterminate 扫动条动画(看着在动);多表仍按 `i/total` 走离散进度
  - ⑥ 1.2.0 起 `convertImagesInRangeToCell` / `convertCellImagesInRangeToFloat` 在壳侧返 `Promise<number>`(为接遮罩),core 内核仍同步
- **JSON 直渲(P3)** —— 新 prop `:workbook` 接 `WorkbookModel | JsonInput`(优先于 `:src`),绕过 parser 直接构造模型渲染。三种 JsonInput shape 自动识别:① 二维数组(首格 A1) ② 对象数组(首行表头 = keys) ③ `{sheets:[{name,rows,...}]}`。类型自动推断:数字字符串 → number、`TRUE`/`FALSE` → boolean、ISO 日期串 → Date(可关 `:jsonOptions="{ autoInfer: false }"`)。新公开导出 `jsonToWorkbook` / `isWorkbookModel` 给"仅引擎"用户。
- ~~**模板填值(P3)** — `props.template` (`TemplateFillSpec`) + `applyTemplate(spec)`~~ — **已废弃**, 改为本版 1.2.0 顶部的"模板样式 overlay"语义. 见上文 ★ 重设计条目.
- **图片浮动 ⇄ 嵌入 选区批量 + 工具栏入口** —— 新 API `convertImagesInRangeToCell(range)` / `convertCellImagesInRangeToFloat(range, size?)` 把"选区内"批量互转,聚合成单次撤销(复用 `convertImagesToCells` 范式 + 新增 `convert-to-floats` exec kind)。Vue 工具栏新内置 `image-tools` 下拉(选区/整表/整列 浮动 ⇄ 嵌入)发现性翻倍;右键菜单多格选区时直接出现「选区浮动图全部嵌入(N 张) / 选区内嵌图全部浮动化(N 张)」。
- **Cell Inspector(单元格全息体检)** —— 新 API `inspectCell(row,col)`,在 `getCellSnapshot` 之上聚合 **合并区**(`merge`/`isMergeAnchor`)、**覆盖到该格的浮动图**(`floatingImages[]`,从 `sheet.images[].from..to` 反推)、**WPS 内嵌图 DISPIMG**(`cellImage`)、**数据验证范围命中**(`dataValidation`)、**条件格式命中**(`conditional[]`,含规则索引 + 等效样式;`ConditionalEngine.inspectHits` 共享 evaluator,跟渲染层同一份)、超链接、批注。框架无关 `src/core/model/inspect.ts`,headless 流程也可调。
- **长任务进度回调 + AbortSignal 取消** —— 所有耗时操作(`exportImage` / `exportPdf` / `print` / `exportXlsx`)统一接 `onProgress?: (p: ExportProgress) => void` + `signal?: AbortSignal`。串行多表导出 + 每页/每表 emit `{stage,sheetIndex,ratio,label}` + `await yieldToEvent()` 让出 UI(防假死)。任意时刻 `ctrl.abort()` → 下一个调度点抛 `DOMException('Aborted','AbortError')`(标准语义),上层用 `e.name === 'AbortError'` 区分取消与真错。新 helper `src/core/export/abort.ts`(`yieldToEvent` / `checkAborted` / `isAbortError`)。
- **自动换行(WPS 风格 toggle)**:工具栏内置 `wrap-text` + 右键菜单「自动换行」+ 命令式 API `toggleWrapTextOnSelection()` / `getSelectionWrapState()`(`'all'/'none'/'mixed'`)。行为对齐 WPS:选区**全已换行 → 全部关掉**;否则**全部打开**;mixed → all。行高自动按内容重撑(失效 autofit 缓存触发),延续"只扩不缩"语义。入命令栈单次撤销 style。两壳同构(Vue 工具栏 `wrap-text` builtin + 右键;React 通过命令式 API + 右键)。
- **HiDPI/系统缩放 canvas 对齐**:Windows 125%/150% 缩放、浏览器 Ctrl+缩放(`devicePixelRatio≠1`)下,canvas 作为"替换元素"默认会按 `width*dpr` 显示,导致整个网格被放大、与 DOM 叠加层(浮动图/图表/HTML 文本框)及鼠标命中错位(越往右下偏得越多)。修复:`render()` 显式把 canvas 的 CSS `width/height` 钉成 view 逻辑尺寸,缓冲仍是 dpr 倍(高清),被浏览器降采样显示 → 像素与逻辑坐标 1:1 对齐。
- **富粘贴(从 Excel/WPS 复制整块)**:`Ctrl+V` 现在优先读剪贴板 **text/html**,完美解析**字体/颜色/填充/边框/对齐/合并单元格**(与 WPS 一致),整块**单次撤销**(整簿快照逆)。值优先取 Excel 的 `x:num`/`x:fmla`(原始值),否则取文本交类型推断。无 html 时回退原 TSV(`pasteText`,不变)。新 `clipboard-html.ts parseClipboardHtml`;`EditController.pasteRich`;API `pasteRichHtml(html,at?)`。
  - **图片(多通道,硬需求)**:① HTML 里的 data-uri `<img>` → 落格转内嵌图;② 剪贴板单张图片(`image/png` 等)→ 落活动格(`pasteImageBlob`);③ 拖图片文件进网格(消费方接 `pasteImageBlob`)。**已知边界**:Excel/WPS **区域复制的内嵌图一般进不了浏览器剪贴板**(只给 text/html+text/plain),所以"复制一整块带图、粘贴时图一起来"做不到 —— 这是浏览器固有限制,需用上述替代通道(单独复制一张图 / 拖文件)。
- **图片点击放大 + 下载原图**:网格里的图(WPS 内嵌图 DISPIMG / 浮动图)可点开看大图、下载原始字节。框架无关 `LightboxHost`(body 级暗背景 + 居中大图 + 「下载原图」+ 点背景/Esc/关闭按钮)。触发:**只读模式单击图**放大;**编辑模式右键**「查看大图 / 下载原图」(不抢选区/编辑)。新 prop `imageLightbox`(默认 true);新 API `openImageLightbox(src,fileName?)` / `getCellImageAt(row,col)`。顺带给点击加 3px 拖拽死区(微抖不再被当拖动,单击语义更稳)。修了 Vue 布尔 prop 缺省被判 false 的坑(`imageLightbox` 加进 withDefaults)。
- **虚拟空行/空列(滚动自动延伸,不动 dimension)**:滚到数据末尾下方仍有空行/空列可滚动、选中、编辑,像 Excel/WPS 的"无限网格";但**不写进 dimension/文件**(避免体积变大)——只有真去编辑某空格,它才靠 `growDimension` 变实。`GridMetrics` 加虚拟范围(`vRows/vCols/virtualWidth/Height`,封顶 Excel 上限 1048576×16384);`totalWidth/Height` 仍按 dimension(**导出/data-access 不含虚拟空行**);spacer 尺寸 / 可视区 / 命中夹取改用虚拟范围;控制器 `recomputeVirtualExtent()`(滚动/缩放/resize 时只增不减、按需延伸)。纯 core,双壳自动继承;新 API `getVirtualExtent()`。
- **背景色 / 字体色(回显 + 修改)**:新 API `getActiveFillColor()` / `getActiveFontColor()`(回显活动格当前色,#RRGGBB)+ `setSelectionFill(color|null)` / `setSelectionFontColor(color)`(改选区,入命令栈)。两壳 demo 加 WPS 风格的背景/字体取色器 + 清除填充。
- **修复内嵌图"灰底"**:DISPIMG 图加载中 / 缺登记项时,之前画 `#f2f4f7` 灰底盖住了单元格自身填充色(白),看着像默认灰。改为加载中不画底色(露出单元格白填充),仅缺图时画个淡图标、不盖色。
- **公式栏(Fx 内容条)可编辑 + 联动**:顶部公式栏从只读改为可编辑 `<input>`(editable + 该格非只读时)。在栏里输入提交(回车下移、Esc 取消、失焦提交)→ 改活动格;切选区 / 格内编辑 → 栏即时反映。栏显示**可编辑字符串**:公式 `=...`、数值原始数字串(非格式化,避免编辑货币/千分位被当文本)、布尔 TRUE/FALSE。新 API:`getCellEditString()` / `canEditActiveCell()` / `commitActiveCellValue(value, move?)`。仅值真变化才入命令栈。两壳同构(React 顺带补:`cell-change` 触发 chrome 重渲)。
- **合并/拆分单元格**:`mergeCells(range)`(吸收相交旧合并、清空被覆盖格只留左上锚点)/ `unmergeCells(range)`,入命令栈可撤销。
- **粘贴**:`Ctrl+V` / `pasteText(text, at?)` —— TSV(制表符+换行)→ 区域写入,类型自动推断、跳过只读、整块一次撤销。
- **右键上下文菜单**:框架无关 body 级 DOM 菜单 —— 插入/删除行列、合并/拆分、清除内容、复制/粘贴;点外部/Esc/滚动关闭、贴边翻转。只读仍用浏览器默认菜单。
- **WPS 单元格内嵌图(DISPIMG)**:
  - **展示**:解析 WPS 私有件 `xl/cellimages.xml` + 单元格 `=DISPIMG("id",1)` 公式,把图按行高列宽画进单元格内(随网格滚动/裁剪/冻结/缩放),普通工具打不开的 WPS 内嵌图现在能正常显示。新 `parser/cell-image-parser.ts`;模型加 `WorkbookModel.cellImages` + `CellModel.dispImgId`;canvas 渲染带图片解码缓存 + onload 重绘。
  - **贴合方式可配置** `cellImageFit`:`contain`(默认,等比缩放,与 WPS 渲染一致——WPS 打开导出文件时 DISPIMG 固定按 contain 显示)/ `fill`(拉伸铺满随格变形)/ `cover`(等比裁剪铺满),两壳 prop + `setCellImageFit()` 运行时切换、即时重绘。
  - **行 customHeight 保真**:解析 `<row customHeight="1">` 标记(ExcelJS 不暴露),自动行高跳过手动设高的行——避免长文本把"作者设矮放图"的行撑大,渲染/导出行高都与 WPS 一致。顺带放宽 fast-xml-parser 实体展开上限(大表几百个 `&quot;` 会撞默认 1000 上限致 drawing/row-meta 解析静默失败)。
  - **就近 / 批量嵌入**:`convertImageToCellAuto(imgIdx)`(图压在哪格就嵌哪格,几何反推)/ `convertAllImagesToCells(col?)`(整表或整列批量,一次进撤销栈)/ `convertCellImageToFloat(row,col,size?)`(嵌入→浮动);右键单格菜单为「将此处浮动图嵌入 / 整列嵌入(N 张)/ 整表嵌入(N 张)/ 内嵌图转浮动图」。两向入命令栈(整簿快照逆)、发 `cell-change`/`image-change`、翻脏标记。`getCellImages()` 读登记表。`convertImageToCell(imgIdx,row,col)` 保留(显式目标格)。
  - **导出往返**:ExcelJS 写出后在 zip 层回注 `cellimages.xml` + rels + media + `[Content_Types].xml`/`workbook.xml.rels` 补丁(新 `export/wps-cellimages.ts`,从模型重建),rebuild / overlay 两模式均覆盖。原有的 + App 内新转的内嵌图导出后用 WPS 打开都显示。验证:解析→导出→再解析 往返存活(单测)。
  - **逐字节对齐真·WPS(修正 #REF!)**:首版用了 `cellimages+xml`(复数)内容类型 + `2017/etCustomData` 关系类型 + 空 `<xdr:spPr/>`,导致 WPS 加载不了登记表、DISPIMG 显示 `#REF!`。据真·WPS 文件实测修正为 **单数** `cellimage+xml`、关系类型 **`2020/cellImage`**、`<xdr:spPr>` 补全 `xfrm`+`prstGeom rect`、`cNvPr` 加 `descr`、DISPIMG 格缓存值 `<v>` 写为 `=DISPIMG("id",1)`。单测锁定这些字段。

### 性能
- **undo 快照轻量化**:增删行列的整簿快照从 `structuredClone`(深拷图片字节/图表)改为手写轻量克隆 —— 只克隆编辑会动的部分,**图片字节/图表/形状/条件格式按引用共享**(编辑期间不可变),大文件 + undo 栈内存大幅下降。惠及结构编辑 undo、脏状态 baseline、还原原件。

### 修复
- **导出失败不再静默**:React 壳「导出 PDF」之前只把错误转给 `onError` prop,未接就被吞(无反应、控制台无报错)。两壳统一为 `console.error` + 上报(emit/onError)+ alert。**提醒**:PDF 导出需可选 peer `jspdf`(`npm i jspdf`);未装时现在会明确报错。

## [1.1.0] - 2026-06-05

把 1.0.0 编辑能力的三处已知 v1 限制做成增强(向后兼容)。

### 公式引用自动重写(增删行列)
- 增删行/列后,自动重写全簿公式里指向该表的 A1 引用(`=A5` 插一行→`=A6`;删被引用行→`#REF!`),
  含绝对/相对 `$`、跨表 `Sheet1!A5`/`'My Sheet'!A5`、区域(删除收缩、全删 `#REF!`);跳过字符串字面量
  与函数名。结构命令改为整簿快照(`cloneWorkbook`)→ 跨表重写也可撤销;开 `recalc` 时引擎按新公式重建。
- 新 `formula/refs.ts`:`shiftFormulaRefs` / `rewriteWorkbookFormulas`。

### 图片导出保真
- 区分锚型:有 `to` 的双格锚 → ExcelJS `br`(随单元格缩放);单格锚 → `tl`(含子格 EMU 偏移转分数列/行)
  + 像素 `ext`;`editAs` 跟随模型。不再一律导成 oneCell+ext、不再丢子格偏移。

### .xlsx 高保真 overlay 导出
- 新 `exportXlsx({ fidelity: 'overlay' })`:重载原始 .xlsx,只把编辑后的 值/样式/合并/行高列宽/冻结
  叠加上去,**保留** ExcelJS 能往返的其余部分(条件格式/数据验证/打印设置/定义名/图表 等)——默认 `rebuild`
  会丢这些。组件加载时留存原件字节供其使用,缺原件自动回退 `rebuild`。overlay 不反映 增删行列/图片 编辑。

### 测试
- 188 单测 + 40 e2e + build 全绿;core 仍零 vue/react/hyperformula/exceljs 静态 import。

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
- **React 壳**:`ooxml-excel-editor/react` 导出 `<ExcelViewer>`(`forwardRef` + 命令式 `ExcelViewerHandle`)与 `useExcelDocument`,与 Vue 壳**共用 ~100% core 引擎**。
- **框架无关 core 入口**:`ooxml-excel-editor/core` 暴露引擎(`ViewerController` / `CanvasRenderer` / `WorkbookExporter` / `OverlayManager` / `PluginOverlayHost`)+ 解析 + 读数据 + 类型,零框架依赖。
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
