// 包根兼容入口 — 老打包器(webpack 4 / vue-cli 4)不读 package.json#exports,
// 解析 `ooxml-excel-editor/vue2` 时走包根文件路径, 找到本文件后 re-export dist 产物.
// webpack 5 / Vite / Rollup 仍走 package.json#exports → ./dist/vue2.js (这个 stub 不会用到).
export { default } from './dist/vue2.js'
export * from './dist/vue2.js'
