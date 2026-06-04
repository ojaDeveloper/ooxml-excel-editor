import { defineConfig } from '@playwright/test'

/**
 * 真浏览器 e2e —— 覆盖 node 单测测不了的 canvas/jsPDF 真实绘制与下载全链路。
 * 用 `npm run test:e2e` 跑(不在 `npm test` 里,保持单测快)。
 * 启动 vite dev(端口 5300),无头 Chromium 驱动 UI: 加载示例 → 渲染 → 导出 → 校验产物。
 */
export default defineConfig({
  testDir: './e2e',
  testMatch: '**/*.e2e.ts',
  fullyParallel: false,
  workers: 1,
  timeout: 60_000,
  expect: { timeout: 15_000 },
  reporter: [['list']],
  use: {
    baseURL: 'http://localhost:5300',
    browserName: 'chromium',
    headless: true,
    acceptDownloads: true,
    viewport: { width: 1280, height: 800 },
  },
  webServer: {
    command: 'npm run dev',
    port: 5300,
    reuseExistingServer: true,
    timeout: 120_000,
  },
})
