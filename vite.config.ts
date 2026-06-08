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
  const isLibBuild = command === 'build' && !isDemoSite && !isVue2Build
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

  return {
    base: isDemoSite ? './' : '/',
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
          lib: {
            entry: { vue2: fileURLToPath(new URL('./src/vue2/index.ts', import.meta.url)) } as Record<string, string>,
            formats: ['es'],
          },
          rollupOptions: {
            external: ['vue', 'vue2', '@vue/composition-api', 'react', 'react-dom', 'react/jsx-runtime', 'exceljs', 'echarts', 'jspdf', 'hyperformula'],
            output: {
              entryFileNames: '[name].js',
              chunkFileNames: 'chunks/[name]-[hash].js',
              assetFileNames: (info) =>
                info.name && info.name.endsWith('.css') ? 'vue2.css' : 'assets/[name][extname]',
              // 'vue2' 别名 → 消费者的 'vue';@vue/composition-api 保持不变, 消费者自己解析
              paths: { vue2: 'vue' },
            },
          },
          chunkSizeWarningLimit: 1500,
        }
      : {
          copyPublicDir: false,
          lib: {
            entry: {
              core: fileURLToPath(new URL('./src/core/index.ts', import.meta.url)),
              index: fileURLToPath(new URL('./src/index.ts', import.meta.url)),
              react: fileURLToPath(new URL('./src/react/index.ts', import.meta.url)),
            },
            formats: ['es'],
          },
          rollupOptions: {
            external: ['vue', 'react', 'react-dom', 'react/jsx-runtime', 'exceljs', 'echarts', 'jspdf', 'hyperformula'],
            output: {
              entryFileNames: '[name].js',
              chunkFileNames: 'chunks/[name]-[hash].js',
              assetFileNames: (info) =>
                info.name && info.name.endsWith('.css') ? 'style.css' : 'assets/[name][extname]',
            },
          },
          chunkSizeWarningLimit: 1500,
        },
  }
})
