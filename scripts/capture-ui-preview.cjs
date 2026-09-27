const { _electron: electron } = require('@playwright/test');
const { mkdtemp, rm } = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');

async function main() {
  const root = path.resolve(__dirname, '..');
  const userData = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-preview-'));
  let application;
  try {
    application = await electron.launch({
      args: [root],
      cwd: root,
      env: { ...process.env, REPORTER_USER_DATA_DIR: userData, REPORTER_TEST_TRAY: '1' },
    });
    const window = await application.firstWindow({ timeout: 15_000 });
    const pages = [
      ['總覽', '今日監控總覽', 'overview'],
      ['關注公司', '關注公司', 'watchlist'],
      ['重大訊息', '重大訊息', 'material'],
      ['違約交割', '全市場違約交割', 'disclosures'],
      ['法說會', '法說會與 Google Calendar', 'conferences'],
      ['設定', '監控排程', 'settings'],
    ];
    for (const [tab, heading, filename] of pages) {
      await window.getByRole('tab', { name: tab }).click();
      await window.getByRole('heading', { name: heading }).waitFor();
      const output = path.join(root, 'artifacts', `ui-preview-${filename}.png`);
      await window.screenshot({ path: output });
      process.stdout.write(`${output}\n`);
    }
  } finally {
    if (application) {
      await application.evaluate(({ app }) => app.exit(0)).catch(() => undefined);
      await application.close().catch(() => undefined);
    }
    if (userData.startsWith(`${os.tmpdir()}${path.sep}`)) {
      await rm(userData, { recursive: true, force: true });
    }
  }
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`);
  process.exitCode = 1;
});
