# 贡献指南(CONTRIBUTING)

读图先看 [ARCHITECTURE](./ARCHITECTURE.md)。这里是怎么跑起来、怎么改、怎么保证不破。

## 环境

- Node ≥ 18,包管理用 `npm`(仓库带 `package-lock.json`)。
- 真浏览器 e2e 需要 Chromium:`npx playwright install chromium`(`@playwright/test` 固定 `1.58.0`,对应缓存的 chromium-1208,避开需联网下载的新版)。

```bash
npm install
npx playwright install chromium   # 首次跑 e2e 前
```

## 常用命令

```bash
npm run dev          # 本地 demo,Vue: http://localhost:5300/   React: http://localhost:5300/react.html
npm test             # 单元测试(vitest,node)
npm run test:e2e     # 真浏览器 e2e(Playwright)
npm run typecheck    # vue-tsc --noEmit(同时检查 .vue / .ts / .tsx)
npm run build        # 构建库:dist/ 三入口 core.js + index.js + react.js + style.css + .d.ts
node scripts/gen-sample.mjs   # 重新生成 public/sample.xlsx(e2e 的"加载示例"用)
```

## 改动前必读:不可破坏的硬约束

1. **测试是回归网**。提交前四条全绿:
   ```bash
   npm run typecheck && npm test && npm run test:e2e && npm run build
   ```
   当前基线:**188 单测 + 40 e2e(Vue + React 双覆盖)**。新功能要补测试。
2. **core 不依赖框架**:`src/core/**` 不得 `from 'vue'` / `'react'`;构建后 `dist/core.js` 同样不得 import 框架(及 hyperformula/exceljs —— 重依赖全动态懒加载)。
3. **两壳同构**:给 `ViewerController` 加能力,Vue 壳(`src/components/ExcelViewer.vue`)与 React 壳(`src/react/ExcelViewer.tsx`)都要接上,e2e 各自覆盖。
4. **默认只读、零回归**:`editable` 关闭时行为与历史一致;编辑能力(值/样式/结构/图片/公式重算/导出回写)是 opt-in。
5. **不打包 peer**:`vue`/`react`/`react-dom`/`exceljs`/`echarts`/`jspdf` 始终 external,产物里不许出现 exceljs(1.4MB)。改构建后用 `npm pack --dry-run` 核对。

## 典型改动流程

1. 在 `src/core/` 实现能力(逻辑/渲染/导出)。先写/补单测。
2. 在 `ViewerController` 暴露命令式方法 +(如需)新增 hook。
3. Vue 壳接上:`onMounted` 里挂 hook(`xxxVersion.value++`),模板/计算属性消费。
4. React 壳接上:hook 里 `force()`;涉及绘制/重建的副作用用 `useLayoutEffect`。
5. 两套各补一条 e2e(Vue → `e2e/*.e2e.ts` 走 `/`;React → `e2e/react.e2e.ts` 走 `/react.html`)。
6. 改了公开 API / 扩展点 / 导出选项 → **同步更新 README 对应表格 + CHANGELOG**。
7. 跑四条全绿 → 提交。

### e2e 小贴士

- 就绪判定要等**真正绘制完成**:画布像素检查须要求**不透明(alpha=255)且非纯白**的像素 —— 否则未绘制的透明画布(r=0)会假阳性。
- demo 在 `import.meta.env.DEV` 下把命令式句柄挂到 `window.__excelViewer`(Vue)/ `window.__excelViewerReact`(React),e2e 可借此算几何 / 读数据。

## 提交信息

中文简述 + 类型前缀(`feat`/`fix`/`refactor`/`docs`/`build`)。一次提交对应一个绿色检查点。

## 目录速查

```
src/core/        框架无关引擎(见 ARCHITECTURE)
src/components/  Vue 壳:ExcelViewer.vue + 子 SFC(toolbar/find/filter/dialog/tabs)
src/composables/ Vue 加载 hook + worker-client(解析 worker;库构建时被 stub 顶替)
src/react/       React 壳:ExcelViewer.tsx + use-excel-document
src/index.ts     Vue 入口   src/react/index.ts React 入口   src/core/index.ts core 入口
src/App.vue / main.ts        Vue demo      react.html / src/react-demo/  React demo
e2e/             Playwright 真浏览器测试
scripts/         gen-sample 等
```
