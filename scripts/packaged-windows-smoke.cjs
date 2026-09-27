const assert = require('node:assert/strict');
const { access, mkdtemp, readFile, rm } = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { _electron: electron } = require('@playwright/test');

async function main() {
  const executable = process.env.REPORTER_PACKAGED_EXECUTABLE
    ? path.resolve(process.env.REPORTER_PACKAGED_EXECUTABLE)
    : path.resolve(process.cwd(), 'artifacts/forge-out/股市記者小幫手-win32-x64/股市記者小幫手.exe');
  await access(executable).catch(() => { throw new Error('找不到 packaged app；請先執行 npm run build'); });
  const isolatedRoot = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-packaged-smoke-'));
  const userData = path.join(isolatedRoot, 'userData');
  const app = await electron.launch({
    executablePath: executable,
    args: [],
    env: {
      ...process.env,
      REPORTER_USER_DATA_DIR: userData,
      ...(process.env.REPORTER_TEST_NO_NODE === '1'
        ? { Path: `${process.env.SystemRoot ?? 'C:\\Windows'}\\System32;${process.env.SystemRoot ?? 'C:\\Windows'}` }
        : {}),
    },
  });
  try {
    const window = await app.firstWindow({ timeout: 15_000 });
    await window.getByRole('heading', { name: '股市記者小幫手' }).waitFor({ state: 'visible', timeout: 10_000 });
    await window.getByRole('tab', { name: '設定' }).click();
    await window.getByRole('heading', { name: '監控排程' }).waitFor({ state: 'visible' });
    await window.getByRole('combobox', { name: '重大訊息檢查頻率' }).waitFor({ state: 'visible' });
    await window.getByRole('checkbox', { name: '收到本日無違約揭露通知' }).waitFor({ state: 'visible' });
    await window.getByRole('button', { name: '匯出資料' }).waitFor({ state: 'visible' });
  } finally {
    await app.close();
    const log = await readFile(path.join(userData, 'logs', 'application.log'), 'utf8').catch(() => '');
    assert.match(log, /"event":"renderer-loaded"/, 'packaged renderer should finish loading');
    assert.doesNotMatch(log, /launch-failed|"event":"renderer-process-gone"/, 'packaged renderer should stay alive');
    await rm(isolatedRoot, { recursive: true, force: true });
  }
  process.stdout.write('Packaged Windows app smoke passed with isolated temporary userData.\n');
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`);
  process.exitCode = 1;
});
