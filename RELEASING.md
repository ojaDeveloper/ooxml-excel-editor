# 发布前 Checklist(待办,尚未发布)

组件库已"产品化就绪",但**还没执行 `npm publish`**。真要发布时按此清单走:

## 1. 改占位信息

[package.json](./package.json) 里这几项现在是占位,发布前改成真实值:

- `author`:目前为空 → 填你的名字/邮箱
- `repository.url` / `homepage` / `bugs.url`:目前是 `your-org/ooxml-excel-preview` → 改成真实 GitHub 仓库
- [LICENSE](./LICENSE) 的版权人:`ooxml-excel-preview contributors` → 可改成你的名字

## 2. 包名可用性

`ooxml-excel-preview` 如果在 npm 已被占用:
- 改名,或加 scope:`@你的用户名/ooxml-excel-preview`(scope 包发布要加 `--access public`)

## 3. 发布

```bash
npm login
npm run build          # prepublishOnly 也会自动跑一次
npm pack --dry-run     # 最后确认产物(应 ~61KB,不含 exceljs)
npm publish            # scope 包: npm publish --access public
```

## 4. CI / 在线 demo(需先推到 GitHub)

- [.github/workflows/ci.yml](./.github/workflows/ci.yml):推送/PR 自动 typecheck+test+build
- [.github/workflows/deploy-demo.yml](./.github/workflows/deploy-demo.yml):推送 main 自动部署 demo 到 GitHub Pages
  - 需在仓库 Settings → Pages → Source 选 "GitHub Actions"
  - `public/sample.xlsx` 要一并提交(demo 的"加载示例"用)

## 已确认的产物边界

- 发布包:`dist/`(ESM 库 155KB + `style.css` + 完整 `.d.ts`),**不含** exceljs/echarts/vue(peer deps)
- 解析:库走主线程;dev/demo 用 Web Worker(见 README "浏览器支持")
