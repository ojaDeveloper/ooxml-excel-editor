# 发布 Checklist(尚未发布)

库已"产品化就绪",但**还没执行 `npm publish`**。真要发布时按此清单走。

## 1. 填占位元数据(发布前必改)

[package.json](./package.json) 这几项发布前改成真实值:

- `author`:目前为空 → 你的名字 / 邮箱
- `repository.url` / `homepage` / `bugs.url`:目前是 `your-org/ooxml-excel-editor` → 真实 GitHub 仓库
- [LICENSE](./LICENSE) 版权人:可改成你的名字

> 这些是**对外发布的身份信息**,需你本人确认;占位值发布出去会留下死链。

## 2. 包名可用性

```bash
npm view ooxml-excel-editor version   # 报 404 = 名字可用;有版本号 = 已被占
```
被占就改名,或加 scope:`@你的用户名/ooxml-excel-editor`(scope 包发布要 `--access public`)。

## 3. 发布

```bash
npm login
npm run build          # prepublishOnly 也会自动跑一次(先 vue-tsc 再三入口 vite build)
npm pack --dry-run     # 最后确认产物
npm publish            # scope 包: npm publish --access public
```

发布产物(`files: ["dist"]`):

| 文件 | 内容 |
|---|---|
| `dist/core.js` (~220KB) | 框架无关引擎 |
| `dist/index.js` (~50KB) | Vue 壳(import `./core.js`) |
| `dist/react.js` (~17KB) | React 壳(import `./core.js`) |
| `dist/style.css` | Vue 壳样式 |
| `dist/*.d.ts` + 类型树 | 三入口 `core.d.ts` / `index.d.ts` / `react.d.ts` |

**不含** `exceljs`/`echarts`/`vue`/`react`/`jspdf`(全 peer);exceljs 仅运行时 `import()`。校验:
```bash
grep -l "exceljs.min" dist/*.js   # 应无输出(没把 1.4MB 库打进去)
```

## 4. CI / 在线 demo(需先推到 GitHub)

- [.github/workflows/ci.yml](./.github/workflows/ci.yml):推送/PR 自动 typecheck+test+build
- [.github/workflows/deploy-demo.yml](./.github/workflows/deploy-demo.yml):推送 main 自动部署 demo
  - 仓库 Settings → Pages → Source 选 "GitHub Actions"
  - `public/sample.xlsx` 一并提交(demo "加载示例"用)

## 5. 版本与变更

- 按 [CHANGELOG.md](./CHANGELOG.md) 把 `[Unreleased]` 收成具体版本号 + 日期。
- 语义化版本:破坏性改 major、加功能 minor、修 bug patch。
