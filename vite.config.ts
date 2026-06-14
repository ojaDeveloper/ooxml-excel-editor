import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import vue2 from '@vitejs/plugin-vue2'
import react from '@vitejs/plugin-react'
import dts from 'vite-plugin-dts'
import { fileURLToPath, URL } from 'node:url'

/**
 * 构建 / 开发模式矩阵
 *
 * 开发 (vite dev):
 *   `npm run dev`        默认 = Vue 3 demo,  port 5300, plugin-vue
 *   `npm run dev:vue3`   alias
 *   `npm run dev:react`  mode=react,         port 5301, plugin-react + plugin-vue (避免 .vue 文件解析失败)
 *   `npm run dev:vue2`   mode=vue2,          port 5302, plugin-vue2 (root=vue2-demo/)
 *
 * 构建 (vite build):
 *   `npm run build`      默认 = 组件库 Vue 3 + React 入口 (lib)
 *   `npm run build:vue2` mode=lib-vue2 (Vue 2 入口, 单独走, append 到 dist/)
 *   `npm run build:demo` mode=demo (站点)
 */
export default defineConfig(({ mode, command }) => {
  const isDemoSite = mode === 'demo'
  const isVue2Build = mode === 'lib-vue2'
  // 库构建(含 Vue 2 入口) 都把 worker-client 别名成 stub(纯主线程)
  // → 不预打包 worker → 产物不含 `new Worker(new URL(..., import.meta.url))` (老 webpack/老 vue-cli 不识别),
  //   也不嵌入 1.4MB exceljs 进 worker chunk. 消费方可自己用 parseWorkbook + 自己包 Worker.
  const isLibBuild = command === 'build' && !isDemoSite
  const isDev = command === 'serve'
  const devTarget: 'vue3' | 'react' | 'vue2' | null = isDev
    ? mode === 'react' ? 'react' : mode === 'vue2' ? 'vue2' : 'vue3'
    : null

  // Vue 2 build / Vue 2 dev 用 plugin-vue2 (隔离 SFC 编译器避免跟 vue@3 冲突)
  // 其它情况用 plugin-vue + plugin-react
  const usePluginVue2 = isVue2Build || devTarget === 'vue2'
  const plugins = [
    ...(usePluginVue2 ? [vue2()] : [vue(), react()]),
    ...(isLibBuild
      ? [dts({
          include: ['src/**/*.ts', 'src/**/*.tsx', 'src/**/*.vue'],
          exclude: ['src/**/__tests__/**', 'src/main.ts', 'src/App.vue', 'src/env.d.ts', 'src/react-demo/**', 'src/vue2/**'],
          insertTypesEntry: true,
          tsconfigPath: './tsconfig.json',
        })]
      : []),
  ]

  const port = devTarget === 'react' ? 5301 : devTarget === 'vue2' ? 5302 : 5300
  // Vue 2 dev 把 root 切到 vue2-demo/, 跟 Vue 3 完全隔离 (不同 plugin 不同进程不同端口)
  // 但是 publicDir 仍指向项目根 public/, 这样 /sample.xlsx 能正常加载
  const rootDir = devTarget === 'vue2'
    ? fileURLToPath(new URL('./vue2-demo', import.meta.url))
    : undefined

  // 反代/内网穿透挂到子路径时 (例 https://host:port/ooxml-excel/), 给 dev server 配 base 前缀,
  // 否则 index.html 里的 /src/main.ts /@vite/client 绝对路径打不到反代的子路径 location → 回退 HTML → MIME 报错.
  // 用法: DEV_BASE=/ooxml-excel/ DEV_PUBLIC_HOST=frp-cat.com DEV_PUBLIC_PORT=59400 npm run dev
  // (HMR WebSocket 也要走同样的对外 host/port + wss, 否则热更连不上.)
  const devBase = isDev && process.env.DEV_BASE ? process.env.DEV_BASE : '/'
  const publicHost = process.env.DEV_PUBLIC_HOST
  const hmr = isDev && publicHost
    ? { protocol: 'wss' as const, host: publicHost, clientPort: Number(process.env.DEV_PUBLIC_PORT) || 443 }
    : undefined

  return {
    base: isDemoSite ? './' : devBase,
    root: rootDir,
    publicDir: devTarget === 'vue2'
      ? fileURLToPath(new URL('./public', import.meta.url))
      : 'public',
    worker: { format: 'es' },
    plugins,
    server: {
      host: true,
      port,
      strictPort: true,
      // 放行经内网穿透/反代域名 (frp-cat.com 等) 访问 dev server —— Vite 默认只放行 localhost.
      // 仅本地开发用; 不影响 lib/demo 构建产物.
      allowedHosts: true,
      ...(hmr ? { hmr } : {}),
      open: devTarget === 'react' ? '/react.html' : '/',
    },
    resolve: {
      alias: [
        ...(isLibBuild
          ? [{
              find: /^(\.\/|@\/composables\/)worker-client$/,
              replacement: fileURLToPath(new URL('./src/composables/worker-client.stub.ts', import.meta.url)),
            }]
          : []),
        // 'vue2' alias 显式指到 compiler-included build (vue.esm.js)
        // lib build 时 rollup output.paths { vue2: 'vue' } 把 'vue2' 重写成 'vue'
        { find: /^vue2$/, replacement: fileURLToPath(new URL('./node_modules/vue2/dist/vue.esm.js', import.meta.url)) },
        // '@vue/composition-api' dev 时重定向到 vue@2.7 dist (内置 Composition API),
        // build 时 rollup external — 消费者: Vue 2.6 走 @vue/composition-api plugin, Vue 2.7+ noop wrapper
        { find: /^@vue\/composition-api$/, replacement: fileURLToPath(new URL('./node_modules/vue2/dist/vue.esm.js', import.meta.url)) },
        { find: '@', replacement: fileURLToPath(new URL('./src', import.meta.url)) },
      ],
    },
    optimizeDeps: {
      include: ['exceljs', 'echarts', 'fflate', 'fast-xml-parser'],
    },
    build: isDemoSite
      ? { chunkSizeWarningLimit: 1500 }
      : isVue2Build
      ? {
          emptyOutDir: false,
          copyPublicDir: false,
          outDir: fileURLToPath(new URL('./dist', import.meta.url)),
          // ES2017 target: 把 class fields / `??` / `?.` (ES2020+) 全部降级, 让 webpack 4 /
          // 老 vue-cli 4 既不报 SyntaxError 也不必装 babel transpileDependencies.
          // 跟 ES2018 区别: 把 spread (ES2018) 等也转, hyperformula 的 class fields 必降.
          target: 'es2017',
          lib: {
            entry: { vue2: fileURLToPath(new URL('./src/vue2/index.ts', import.meta.url)) } as Record<string, string>,
            formats: ['es'],
          },
          rollupOptions: {
            // ★ exceljs / jspdf / hyperformula 都 inline 进 chunks/ + ES2017 降级
            // (它们用 class fields / ES2020+ 语法, webpack 4 解析源码会炸).
            // echarts 仍 external — 它现代版 CJS / UMD 入口 webpack 4 兼容良好,
            // 用户项目大概率已有自己的 echarts (避免 dual instance / theme 失效).
            // framework (vue / react / @vue/composition-api) 必 external — 跟宿主同实例.
            external: ['vue', 'vue2', '@vue/composition-api', 'react', 'react-dom', 'react/jsx-runtime', 'echarts'],
            output: {
              entryFileNames: '[name].js',
              chunkFileNames: 'chunks/[name]-[hash].js',
              assetFileNames: (info) =>
                info.name && info.name.endsWith('.css') ? 'vue2.css' : 'assets/[name][extname]',
              paths: { vue2: 'vue' },
            },
          },
          chunkSizeWarningLimit: 5000,
        }
      : {
          copyPublicDir: false,
          // ES2017 target: 把 class fields / `??` / `?.` 全部降级 (一致跟 lib-vue2)
          target: 'es2017',
          lib: {
            entry: {
              core: fileURLToPath(new URL('./src/core/index.ts', import.meta.url)),
              index: fileURLToPath(new URL('./src/index.ts', import.meta.url)),
              react: fileURLToPath(new URL('./src/react/index.ts', import.meta.url)),
            },
            formats: ['es'],
          },
          rollupOptions: {
            // ★ exceljs / jspdf / hyperformula inline 进 chunks/ (老打包器不解析它们源码).
            // echarts 仍 external (用户项目大概率已有, 避免 dual instance / theme 失效).
            external: ['vue', 'react', 'react-dom', 'react/jsx-runtime', 'echarts'],
            output: {
              entryFileNames: '[name].js',
              chunkFileNames: 'chunks/[name]-[hash].js',
              assetFileNames: (info) =>
                info.name && info.name.endsWith('.css') ? 'style.css' : 'assets/[name][extname]',
            },
          },
          chunkSizeWarningLimit: 5000,
        },
  }
})
