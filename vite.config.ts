import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import vue2 from '@vitejs/plugin-vue2'
import react from '@vitejs/plugin-react'
import dts from 'vite-plugin-dts'
import { fileURLToPath, URL } from 'node:url'

/**
 * 三种构建:
 * - 默认 `vite build`            → 组件库 Vue 3 + React 壳 (lib 模式), 产物给别的项目 import
 * - `vite build --mode lib-vue2` → Vue 2 子入口 (单独跑, 跟主构建合在 npm run build 里)
 * - `vite build --mode demo`     → demo 站点(预览/部署用)
 * 开发 `vite dev` 始终走 demo(index.html + main.ts),不受 lib 配置影响。
 */
export default defineConfig(({ mode, command }) => {
  const isDemo = mode === 'demo'
  const isVue2Build = mode === 'lib-vue2'
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
      // Vue 2 入口构建: 只装 plugin-vue2 (处理 .vue, 编译成 Vue 2 runtime), 不要 plugin-vue (会冲突)
      // 其他构建: 装 plugin-vue + plugin-react, 各管各的文件类型
      ...(isVue2Build ? [vue2()] : [vue(), react()]),
      // 仅 lib 构建时生成 .d.ts(只针对 Vue 3 库入口;Vue 2 入口暂不生成 .d.ts, 避免 vue-tsc 不认 Vue 2 SFC)
      ...(isDemo || isVue2Build
        ? []
        : [
            dts({
              include: ['src/**/*.ts', 'src/**/*.tsx', 'src/**/*.vue'],
              exclude: ['src/**/__tests__/**', 'src/main.ts', 'src/App.vue', 'src/env.d.ts', 'src/react-demo/**', 'src/vue2/**'],
              insertTypesEntry: true,
              tsconfigPath: './tsconfig.json',
            }),
          ]),
    ],
    server: {
      // host:true → 监听 0.0.0.0,局域网其它设备可用 http://<本机IP>:5300 访问(dev 启动会打印 Network 地址)。
      host: true,
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
        // 'vue2' alias 显式指到 compiler-included build (vue.esm.js) — dev / build 都生效.
        // 为什么不用默认 module 入口 (vue.runtime.esm.js): demo 里 new Vue({ template: '...' })
        // 需要运行时模板编译器. lib 代码用 h() render function 不用编译器, 多带的编译器会被
        // tree-shake 掉. lib build 时 rollup output.paths { vue2: 'vue' } 把 'vue2' 替换成
        // 'vue', 用户提供自己的 vue@2 peer (用户那边 runtime-only 还是 with-compiler 由用户决定).
        { find: /^vue2$/, replacement: fileURLToPath(new URL('./node_modules/vue2/dist/vue.esm.js', import.meta.url)) },
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
      : isVue2Build
      ? {
          // Vue 2 子入口: 单独跑, 不清 outDir (已有 Vue 3/React 产物)
          emptyOutDir: false,
          copyPublicDir: false,
          lib: {
            entry: { vue2: fileURLToPath(new URL('./src/vue2/index.ts', import.meta.url)) } as Record<string, string>,
            formats: ['es'],
          },
          rollupOptions: {
            // 'vue2' 和 'vue' 都 external — 源码用 'vue2' (开发期 alias 解析到 vue@2.7), 但
            // build 产物里通过 output.paths 把 'vue2' 重写成 'vue', 让消费方用 peer vue@2.7 解析
            external: ['vue', 'vue2', 'react', 'react-dom', 'react/jsx-runtime', 'exceljs', 'echarts', 'jspdf', 'hyperformula'],
            output: {
              entryFileNames: '[name].js',
              chunkFileNames: 'chunks/[name]-[hash].js',
              assetFileNames: (info) =>
                info.name && info.name.endsWith('.css') ? 'vue2.css' : 'assets/[name][extname]',
              paths: { vue2: 'vue' },
            },
          },
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
