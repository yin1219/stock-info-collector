import { defineConfig } from "@playwright/test";
export default defineConfig({
  testDir: "test/demo",
  outputDir: "artifacts/ui-demo/test-results",
  workers: 1,
  timeout: 30000,
  use: {
    channel: process.env.REPORTER_DEMO_BROWSER || "chrome",
    baseURL: "http://127.0.0.1:5178",
    viewport: { width: 1488, height: 1056 },
    trace: "retain-on-failure",
  },
  webServer: {
    command:
      "node node_modules/vite/bin/vite.js --config vite.demo.config.mts --host 127.0.0.1",
    url: "http://127.0.0.1:5178",
    reuseExistingServer: !process.env.CI,
  },
  reporter: "list",
});
