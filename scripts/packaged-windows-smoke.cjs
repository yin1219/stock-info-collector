const assert = require('node:assert/strict');
const { access, mkdtemp, readFile, rm } = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { spawn } = require('node:child_process');
const { _electron: electron } = require('@playwright/test');

async function main() {
  const executable = process.env.REPORTER_PACKAGED_EXECUTABLE
    ? path.resolve(process.env.REPORTER_PACKAGED_EXECUTABLE)
    : path.resolve(process.cwd(), 'artifacts/forge-out/股市記者小幫手-win32-x64/StockReporterAssistant.exe');
  await access(executable).catch(() => { throw new Error('找不到 packaged app；請先執行 npm run build'); });
  const { extractFile } = await import('@electron/asar');
  const notices = extractFile(path.join(path.dirname(executable), 'resources', 'app.asar'), 'assets/third-party-notices.txt').toString('utf8').replace(/\r\n/g, '\n');
  const driverLicense = (await readFile(path.resolve(process.cwd(), 'node_modules/driver.js/license'), 'utf8')).replace(/\r\n/g, '\n');
  assert.ok(notices.includes(driverLicense), 'packaged app must include the full Driver.js MIT notice');

  const isolatedRoot = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-packaged-smoke-'));
  const userData = path.join(isolatedRoot, 'userData');
  const isolatedEnv = {
    ...process.env,
    REPORTER_USER_DATA_DIR: userData,
    ...(process.env.REPORTER_TEST_NO_NODE === '1'
      ? { Path: `${process.env.SystemRoot ?? 'C:\\Windows'}\\System32;${process.env.SystemRoot ?? 'C:\\Windows'}` }
      : {}),
  };
  delete isolatedEnv.REPORTER_TEST_LIVE_SOURCES;
  delete isolatedEnv.REPORTER_TEST_HTTP_PROXY;
  delete isolatedEnv.REPORTER_TEST_LIVE_OAUTH;
  const app = await electron.launch({
    executablePath: executable,
    args: [],
    env: isolatedEnv,
  });
  try {
    const window = await app.firstWindow({ timeout: 15_000 });
    await window.locator('.brand').filter({hasText:'股市記者小幫手'}).waitFor({state:'visible',timeout:10_000});
    await window.getByRole('heading',{name:'總覽',exact:true}).waitFor({state:'visible'});
    const packageVersion = JSON.parse(await readFile(path.resolve(process.cwd(), 'package.json'), 'utf8')).version;
    await window.getByText(`v2 · 版本 ${packageVersion}`, { exact: true }).waitFor({ state: 'visible' });
    await window.getByRole('tab', { name: '設定' }).click();
    await window.getByRole('heading', { name: '設定',exact:true }).waitFor({ state: 'visible' });
    await window.getByRole('combobox', { name: '重大訊息檢查頻率' }).waitFor({ state: 'visible' });
    await window.getByRole('checkbox', { name: '本日無違約揭露時也通知' }).waitFor({ state: 'visible' });
    await window.getByRole('tab',{name:'資料管理',exact:true}).click();
    await window.getByRole('button',{name:'匯出本機資料'}).waitFor({state:'visible'});
    await window.getByRole('tab',{name:'通知',exact:true}).click();
    await window.getByRole('button', { name: '發送測試通知', exact: true }).click();
    await window.getByText('隔離測試模式不顯示 Windows 通知。').waitFor({ state: 'visible' });

    const nativeWindow = await app.browserWindow(window);
    await nativeWindow.evaluate((candidate) => candidate.hide());
    const secondExitCode = await new Promise((resolve, reject) => {
      const child = spawn(executable, ['--conference-scheduled'], { env: isolatedEnv, windowsHide: true, stdio: 'ignore' });
      const timeout = setTimeout(() => { child.kill(); reject(new Error('Second packaged conference entry did not exit')); }, 20_000);
      child.once('error', (error) => { clearTimeout(timeout); reject(error); });
      child.once('exit', (code) => { clearTimeout(timeout); resolve(code); });
    });
    assert.equal(secondExitCode, 0, 'second packaged scheduled process should exit successfully');
    const primaryLogPath = path.join(userData, 'logs', 'application.log');
    const deadline = Date.now() + 10_000;
    let primaryLog = '';
    do {
      primaryLog = await readFile(primaryLogPath, 'utf8');
      if (primaryLog.includes('"event":"conference-scheduled-not-ready"')) break;
      await new Promise((resolve) => setTimeout(resolve, 100));
    } while (Date.now() < deadline);
    assert.match(primaryLog, /"event":"conference-scheduled-not-ready"/, 'scheduled launch should reach the existing packaged process');
    assert.equal((primaryLog.match(/"event":"conference-scheduled-not-ready"/g) ?? []).length, 1, 'one scheduled launch should run once in the existing process');
    assert.equal((primaryLog.match(/"event":"main-ready"/g) ?? []).length, 1, 'scheduled launch must not create another primary process');
    assert.equal((primaryLog.match(/"event":"renderer-load-started"/g) ?? []).length, 1, 'scheduled launch must not open another renderer');
    assert.equal(await nativeWindow.evaluate((candidate) => candidate.isVisible()), false, 'scheduled launch must leave the background window hidden');
    assert.equal(await app.evaluate(({ app: electronApp }) => electronApp.isReady()), true, 'background primary process must remain alive');
  } finally {
    await app.close();
    const log = await readFile(path.join(userData, 'logs', 'application.log'), 'utf8').catch(() => '');
    assert.match(log, /"event":"renderer-loaded"/, 'packaged renderer should finish loading');
    assert.doesNotMatch(log, /launch-failed|"event":"renderer-process-gone"/, 'packaged renderer should stay alive');
    try {
      const scheduledUserData = path.join(isolatedRoot, 'scheduled-userData');
      const headlessEnv = { ...process.env, REPORTER_USER_DATA_DIR: scheduledUserData };
      delete headlessEnv.REPORTER_TEST_LIVE_SOURCES;
      delete headlessEnv.REPORTER_TEST_HTTP_PROXY;
      delete headlessEnv.REPORTER_TEST_LIVE_OAUTH;
      const exitCode = await new Promise((resolve, reject) => {
        const child = spawn(executable, ['--conference-scheduled'], { env: headlessEnv, windowsHide: true, stdio: 'ignore' });
        const timeout = setTimeout(() => { child.kill(); reject(new Error('Packaged conference daily entry did not exit')); }, 20_000);
        child.once('error', (error) => { clearTimeout(timeout); reject(error); });
        child.once('exit', (code) => { clearTimeout(timeout); resolve(code); });
      });
      assert.equal(exitCode, 0, 'isolated conference daily entry should exit successfully');
      const scheduledLog = await readFile(path.join(scheduledUserData, 'logs', 'application.log'), 'utf8');
      assert.match(scheduledLog, /"event":"conference-scheduled-not-ready"/, 'unknown task ownership must fail closed');
      assert.doesNotMatch(scheduledLog, /"event":"renderer-load-started"/, 'daily entry must not open the renderer');
    } finally {
      await rm(isolatedRoot, { recursive: true, force: true });
    }
  }
  process.stdout.write('Packaged Windows app smoke passed with isolated temporary userData.\n');
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`);
  process.exitCode = 1;
});
