import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import react from '@vitejs/plugin-react'
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
      // React 壳(.tsx)走 react 插件;与 vue 插件并存,各管各的文件类型
      react(),
      // 仅 lib 构建时生成 .d.ts(只针对 Vue 库入口;React 壳的 .tsx/react-demo 不进 Vue 包)
      ...(isDemo
        ? []
        : [
            dts({
              include: ['src/**/*.ts', 'src/**/*.tsx', 'src/**/*.vue'],
              exclude: ['src/**/__tests__/**', 'src/main.ts', 'src/App.vue', 'src/env.d.ts', 'src/react-demo/**'],
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
        // 库构建用 stub 顶替 worker-client(去掉 worker → 不打包 exceljs)。
        // 同时匹配 Vue 壳的相对引入(./worker-client)与 React 壳的别名引入(@/composables/worker-client)。
        ...(isLibBuild
          ? [
              {
                find: /^(\.\/|@\/composables\/)worker-client$/,
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
          // 组件库: ESM 多入口 —— core(框架无关引擎) / index(Vue 壳) / react(React 壳)。
          // 同名 chunk 抽到 chunks/ 共享(core 引擎被 vue+react 复用,只打一份)。
          // peer 依赖(vue/react/exceljs/echarts/jspdf)全 external,不打进库。
          copyPublicDir: false, // 不把 public/sample.xlsx 打进库
          lib: {
            entry: {
              core: fileURLToPath(new URL('./src/core/index.ts', import.meta.url)),
              index: fileURLToPath(new URL('./src/index.ts', import.meta.url)),
              react: fileURLToPath(new URL('./src/react/index.ts', import.meta.url)),
            },
            formats: ['es'],
          },
          rollupOptions: {
            // jspdf/hyperformula 同 echarts: 可选 peer,运行时动态 import,不打进库
            external: ['vue', 'react', 'react-dom', 'react/jsx-runtime', 'exceljs', 'echarts', 'jspdf', 'hyperformula'],
            output: {
              entryFileNames: '[name].js',
              chunkFileNames: 'chunks/[name]-[hash].js',
              // 把 css 产物固定命名为 style.css(Vue 壳)/ react.css(React 壳)按需,默认 style.css
              assetFileNames: (info) =>
                info.name && info.name.endsWith('.css') ? 'style.css' : 'assets/[name][extname]',
            },
          },
          chunkSizeWarningLimit: 1500,
        },
  }
})
