import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import dts from 'vite-plugin-dts'
import { fileURLToPath, URL } from 'node:url'

/**
 * 两种构建:
 * - 默认 `vite build`            → 组件库(lib 模式),产物给别的项目 import
 * - `vite build --mode demo`     → demo 站点(预览/部署用)
 * 开发 `vite dev` 始终走 demo(index.html + main.ts),不受 lib 配置影响。
 */
export default defineConfig(({ mode, command }) => {
  const isDemo = mode === 'demo'
  // 库构建(vite build,非 demo): 把 worker-client 别名成 stub(纯主线程),
  // 这样 vite 不会扫到 new Worker(...) → 不预打包 worker → 不把 1.4MB exceljs 打进库。
  // dev / demo 用真正的 worker-client(大文件不卡)。
  const isLibBuild = command === 'build' && !isDemo

  return {
    // demo 用相对 base，部署到 GitHub Pages 子路径也能正确加载资源
    base: isDemo ? './' : '/',
    // 解析 Worker 内部用动态 import(exceljs) → 需 ES 格式(iife 不支持代码分割)
    worker: { format: 'es' },
    plugins: [
      vue(),
      // 仅 lib 构建时生成 .d.ts
      ...(isDemo
        ? []
        : [
            dts({
              include: ['src/**/*.ts', 'src/**/*.vue'],
              exclude: ['src/**/__tests__/**', 'src/main.ts', 'src/App.vue', 'src/env.d.ts'],
              insertTypesEntry: true,
              tsconfigPath: './tsconfig.json',
            }),
          ]),
    ],
    server: {
      // 钉死端口，避免跟其它项目(OKR=5173 / 若依=5174)抢先后顺序。
      // strictPort: 被占就明确报错，而不是默默漂移到别的端口让人找不到。
      port: 5300,
      strictPort: true,
    },
    resolve: {
      alias: [
        // 库构建用 stub 顶替 worker-client(去掉 worker → 不打包 exceljs)
        ...(isLibBuild
          ? [
              {
                find: /^\.\/worker-client$/,
                replacement: fileURLToPath(new URL('./src/composables/worker-client.stub.ts', import.meta.url)),
              },
            ]
          : []),
        { find: '@', replacement: fileURLToPath(new URL('./src', import.meta.url)) },
      ],
    },
    optimizeDeps: {
      // exceljs ships CJS; let Vite pre-bundle it so the browser build works.
      include: ['exceljs', 'echarts', 'fflate', 'fast-xml-parser'],
    },
    build: isDemo
      ? {
          chunkSizeWarningLimit: 1500,
        }
      : {
          // 组件库: ESM 产物。vue / exceljs / echarts 设为 external(peerDependencies)，
          // 不打进库;fflate / fast-xml-parser 体积小且解析核心，打进库。
          copyPublicDir: false, // 不把 public/sample.xlsx 打进库
          lib: {
            entry: fileURLToPath(new URL('./src/index.ts', import.meta.url)),
            name: 'OoxmlExcelPreview',
            formats: ['es'],
            fileName: 'ooxml-excel-preview',
          },
          rollupOptions: {
            // jspdf 同 echarts: 可选 peer,运行时动态 import,不打进库
            external: ['vue', 'exceljs', 'echarts', 'jspdf'],
            output: {
              // 把唯一的 css 产物固定命名为 style.css，方便宿主 import '.../style.css'
              assetFileNames: (info) =>
                info.name && info.name.endsWith('.css') ? 'style.css' : 'assets/[name][extname]',
            },
          },
          chunkSizeWarningLimit: 1500,
        },
  }
})
