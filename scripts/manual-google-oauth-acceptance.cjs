const { _electron: electron, expect } = require('@playwright/test');
const { mkdtemp, rm } = require('node:fs/promises');
const { createInterface } = require('node:readline');
const os = require('node:os');
const path = require('node:path');

async function waitForEnter(message) {
  const input = createInterface({ input: process.stdin, output: process.stdout });
  await new Promise((resolve) => input.question(message, resolve));
  input.close();
}

(async () => {
  const root = path.resolve(__dirname, '..');
  const userData = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-google-oauth-'));
  let app;

  try {
    app = await electron.launch({
      ...(process.env.REPORTER_PACKAGED_EXECUTABLE
        ? { executablePath: path.resolve(process.env.REPORTER_PACKAGED_EXECUTABLE), args: [] }
        : { args: [root] }),
      cwd: root,
      env: {
        ...process.env,
        REPORTER_USER_DATA_DIR: userData,
        REPORTER_TEST_TRAY: '1',
        REPORTER_TEST_LIVE_OAUTH: '1',
      },
    });
    const window = await app.firstWindow({ timeout: 15_000 });
    await window.bringToFront();
    await window.getByRole('tab', { name: '法說會' }).click();
    const initialStatus = (await window.locator('.section-heading .status').textContent())?.trim();
    console.log('GOOGLE_STATUS_BEFORE_LOGIN', initialStatus);
    if (initialStatus !== '尚未連線') {
      throw new Error('應用程式尚未載入 OAuth client 設定；未啟動 Google 登入。');
    }
    await window.getByRole('button', { name: '連線 Google Calendar' }).click();
    await window.waitForTimeout(1_000);
    const connectionError = await window.getByRole('alert').allTextContents();
    if (connectionError.length > 0) throw new Error(connectionError.join('；'));

    console.log('已向 Windows 要求開啟 Google 登入頁；請確認系統瀏覽器實際顯示後自行登入及授權。本程式不會讀取帳密，也不會執行 Calendar 同步。');
    await expect(window.getByText('Google 已連線', { exact: true })).toBeVisible({ timeout: 5 * 60_000 });
    await expect(window.getByRole('button', { name: '立即擷取並同步法說會' })).toBeVisible();
    console.log('OAuth 連線驗收通過；未執行 Calendar 同步。按 Enter 關閉隔離測試視窗並清除暫存授權資料。');
    await waitForEnter('');
  } finally {
    if (app) {
      await app.evaluate(({ app: electronApp }) => electronApp.exit(0)).catch(() => undefined);
      await app.close().catch(() => undefined);
    }
    await rm(userData, { recursive: true, force: true });
  }
})().catch((error) => {
  console.error('OAuth acceptance failed:', error.message);
  process.exitCode = 1;
});
