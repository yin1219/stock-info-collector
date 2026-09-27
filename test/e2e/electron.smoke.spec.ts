import { access, mkdtemp, readFile, rm } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import os from 'node:os';
import path from 'node:path';
import { _electron as electron, expect, test } from '@playwright/test';
import Database from 'better-sqlite3';

test('watchlist-management / Electron IPC / removes a watched company while preserving saved history', async () => {
  const root = path.resolve(process.cwd());
  const userData = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-remove-e2e-'));
  const app = await electron.launch({
    args: [root], cwd: root,
    env: { ...process.env, REPORTER_USER_DATA_DIR: userData, REPORTER_TEST_TRAY: '1' },
  });

  try {
    const window = await app.firstWindow({ timeout: 8_000 });
    const database = new Database(path.join(userData, 'reporter.sqlite3'));
    try {
      database.prepare('INSERT INTO companies (id, market, stock_code, name, updated_at) VALUES (?, ?, ?, ?, ?)')
        .run('company-1', 'TWSE', '2330', '台積電', '2026-09-27T00:00:00Z');
      database.prepare('INSERT INTO watchlist_entries (id, company_id, active, category, notes, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)')
        .run('watch-1', 'company-1', 1, '', '', '2026-09-27T00:00:00Z', '2026-09-27T00:00:00Z');
      database.prepare('INSERT INTO material_events (id, company_id, source_key, content_fingerprint, title, content, published_at) VALUES (?, ?, ?, ?, ?, ?, ?)')
        .run('event-1', 'company-1', 'test-event-1', 'hash', '董事會決議', '公告全文', '2026-09-27T00:00:00Z');
    } finally {
      database.close();
    }

    await window.getByRole('tab', { name: '關注公司' }).click();
    await expect(window.getByText('台積電')).toBeVisible();
    await window.getByRole('button', { name: '移除台積電' }).click();
    await expect(window.getByText('台積電')).toHaveCount(0);
    await expect(window.getByText('目前沒有符合條件的公司。')).toBeVisible();

    const saved = new Database(path.join(userData, 'reporter.sqlite3'), { readonly: true });
    try {
      expect(saved.prepare('SELECT COUNT(*) AS count FROM watchlist_entries').get()).toEqual({ count: 0 });
      expect(saved.prepare('SELECT COUNT(*) AS count FROM material_events').get()).toEqual({ count: 1 });
    } finally {
      saved.close();
    }
  } finally {
    await app.evaluate(({ app: electronApp }) => electronApp.exit(0)).catch(() => undefined);
    await app.close();
    await rm(userData, { recursive: true, force: true });
  }
});

test('material-event-monitoring / Electron UI / defaults to watched announcements and expands details in the clicked row', async () => {
  const root = path.resolve(process.cwd());
  const userData = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-material-e2e-'));
  const app = await electron.launch({
    args: [root], cwd: root,
    env: { ...process.env, REPORTER_USER_DATA_DIR: userData, REPORTER_TEST_TRAY: '1' },
  });
  try {
    const window = await app.firstWindow({ timeout: 8_000 });
    const database = new Database(path.join(userData, 'reporter.sqlite3'));
    try {
      const addCompany = database.prepare('INSERT INTO companies (id, market, stock_code, name, updated_at) VALUES (?, ?, ?, ?, ?)');
      addCompany.run('company-watched', 'TWSE', '2330', '台積電', '2026-09-27T00:00:00Z');
      addCompany.run('company-other', 'TPEX', '6488', '其他公司', '2026-09-27T00:00:00Z');
      database.prepare('INSERT INTO watchlist_entries (id, company_id, active, category, notes, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)')
        .run('watch-1', 'company-watched', 1, '', '', '2026-09-27T00:00:00Z', '2026-09-27T00:00:00Z');
      const addEvent = database.prepare('INSERT INTO material_events (id, company_id, source_key, content_fingerprint, title, content, published_at) VALUES (?, ?, ?, ?, ?, ?, ?)');
      addEvent.run('event-watched', 'company-watched', 'watched-1', 'hash-1', '關注公告', '關注公告全文', '2026-09-27T01:00:00Z');
      addEvent.run('event-other', 'company-other', 'other-1', 'hash-2', '其他公告', '其他公告全文', '2026-09-27T02:00:00Z');
    } finally { database.close(); }
    await window.getByRole('tab', { name: '重大訊息' }).click();
    await expect(window.getByText('關注 1 筆')).toBeVisible();
    await expect(window.getByRole('button', { name: '台積電 關注公告' })).toBeVisible();
    await expect(window.getByRole('button', { name: '其他公司 其他公告' })).toHaveCount(0);
    await window.getByRole('checkbox', { name: '顯示全部公告' }).check();
    await expect(window.getByText('全部 2 筆')).toBeVisible();
    const otherRow = window.getByRole('button', { name: '其他公司 其他公告' }).locator('..');
    await otherRow.getByRole('button', { name: '其他公司 其他公告' }).click();
    await expect(otherRow.getByRole('complementary', { name: '重大訊息完整內容' })).toContainText('其他公告全文');
    const watchedRow = window.getByRole('button', { name: '台積電 關注公告' }).locator('..');
    await watchedRow.getByRole('button', { name: '台積電 關注公告' }).click();
    await expect(watchedRow.getByRole('complementary', { name: '重大訊息完整內容' })).toContainText('關注公告全文');
    await expect(otherRow.getByRole('complementary', { name: '重大訊息完整內容' })).toHaveCount(0);
  } finally {
    await app.evaluate(({ app: electronApp }) => electronApp.exit(0)).catch(() => undefined);
    await app.close();
    await rm(userData, { recursive: true, force: true });
  }
});

test('desktop-app-lifecycle / Electron shell / opens the isolated renderer', async () => {
  const root = path.resolve(process.cwd());
  const userData = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-e2e-'));
  const app = await electron.launch({
    args: [root],
    cwd: root,
    env: {
      ...process.env,
      REPORTER_USER_DATA_DIR: userData,
      REPORTER_TEST_TRAY: '1',
    },
  });

  try {
    const window = await app.firstWindow({ timeout: 8_000 });
    await expect(window.locator('.sidebar')).toBeVisible();
    const layout = await window.evaluate(() => {
      const sidebar = document.querySelector('.sidebar')!;
      const content = document.querySelector('.app-shell > .panel')!;
      return {
        bodyOverflowY: getComputedStyle(document.body).overflowY,
        sidebarPosition: getComputedStyle(sidebar).position,
        sidebarHeight: sidebar.getBoundingClientRect().height,
        viewportHeight: window.innerHeight,
        contentOverflowY: getComputedStyle(content).overflowY,
        navigationVisible: document.querySelector('.main-nav')!.getBoundingClientRect().bottom <= window.innerHeight,
      };
    });
    expect(layout.bodyOverflowY).toBe('hidden');
    expect(layout.sidebarPosition).toBe('sticky');
    expect(layout.sidebarHeight).toBe(layout.viewportHeight);
    expect(layout.contentOverflowY).toBe('auto');
    expect(layout.navigationVisible).toBe(true);
    const navigation = window.getByRole('tablist', { name: '主要功能' });
    await expect.poll(() => navigation.evaluate((element) => getComputedStyle(element).flexDirection)).toBe('column');
    await window.setViewportSize({ width: 600, height: 800 });
    await expect.poll(() => navigation.evaluate((element) => getComputedStyle(element).flexDirection)).toBe('row');
    await window.setViewportSize({ width: 1200, height: 800 });
    window.on('pageerror', (error) => console.error(`renderer:pageerror:${error.message}`));
    await expect(window.getByRole('heading', { name: '股市記者小幫手' })).toBeVisible();
    await expect(window).toHaveTitle('股市記者小幫手 v2');
    await expect.poll(() => window.evaluate(() => typeof window.reporterApi.listWatchlist)).toBe('function');
    await expect.poll(() => window.evaluate(() => typeof window.reporterApi.listDisclosures)).toBe('function');
    await expect.poll(() => window.evaluate(() => typeof window.reporterApi.listMaterialEvents)).toBe('function');
    await expect(window.evaluate(() => window.reporterApi.listWatchlist({ activeOnly: true }))).resolves.toEqual([]);
    await expect(window.evaluate(() => window.reporterApi.listDisclosures({ disclosureDate: '2026-09-26' }))).resolves.toEqual([]);
    await expect(window.evaluate(() => window.reporterApi.listMaterialEvents({ unreadOnly: true }))).resolves.toEqual([]);
    const electronExecutable = require('electron') as string;
    const secondInstance = spawn(electronExecutable, [root], {
      cwd: root,
      windowsHide: true,
      stdio: 'ignore',
      env: { ...process.env, REPORTER_USER_DATA_DIR: userData },
    });
    const secondExitCode = await new Promise<number | null>((resolve, reject) => {
      secondInstance.once('error', reject);
      secondInstance.once('exit', (code) => resolve(code));
    });
    expect(secondExitCode).toBe(0);
    const singleInstanceLog = await readFile(path.join(userData, 'logs', 'application.log'), 'utf8');
    expect(singleInstanceLog.match(/"event":"main-ready"/g)).toHaveLength(1);
    expect(await app.windows()).toHaveLength(1);
    await expect(window.getByRole('heading', { name: '股市記者小幫手' })).toBeVisible();
    await window.getByRole('tab', { name: '違約交割' }).click();
    await expect(window.getByRole('heading', { name: '全市場違約交割' })).toBeVisible();
    await window.getByRole('tab', { name: '重大訊息' }).click();
    await expect(window.getByRole('heading', { name: '重大訊息' })).toBeVisible();
    await window.getByRole('tab', { name: '關注公司' }).click();
    await expect(window.getByRole('heading', { name: '關注公司' })).toBeVisible();
    await window.getByRole('textbox', { name: '股票代號' }).click({ timeout: 2_000 });
    expect(await window.locator('.app-shell > .panel').count()).toBe(1);
    const watchlistLayout = await window.evaluate(() => ({
      addFormBottom: document.querySelector('.add-form')!.getBoundingClientRect().bottom,
      importTop: document.querySelector('.import-panel')!.getBoundingClientRect().top,
    }));
    expect(watchlistLayout.importTop).toBeGreaterThan(watchlistLayout.addFormBottom);
    await window.getByRole('tab', { name: '設定' }).click();
    await expect(window.getByRole('heading', { name: '監控排程' })).toBeVisible();
    await expect(window.getByRole('combobox', { name: '重大訊息檢查頻率' })).toHaveValue('60');
    await expect(window.getByText('隔離測試模式不會連線官方網站；立即檢查已停用。')).toBeVisible();
    await expect(window.getByRole('button', { name: '立即檢查重大訊息' })).toBeDisabled();
    await window.getByRole('tab', { name: '法說會' }).click();
    await expect(window.getByText('尚未設定', { exact: true })).toBeVisible();
    await window.getByRole('button', { name: '連線 Google Calendar' }).click();
    await expect(window.getByRole('alert')).toContainText('Google OAuth 用戶端尚未隨應用程式提供');
    await access(path.join(userData, 'reporter.sqlite3'));
    const applicationLog = await readFile(path.join(userData, 'logs', 'application.log'), 'utf8');
    expect(applicationLog).toContain('"event":"main-ready"');
    expect(applicationLog).toContain('"event":"renderer-loaded"');
    expect(applicationLog).toContain('"event":"google-calendar-action-failed"');
    expect(applicationLog).toContain('"operation":"calendar:connect"');
  } finally {
    const applicationLog = await readFile(path.join(userData, 'logs', 'application.log'), 'utf8').catch(() => '');
    if (applicationLog) console.info(`e2e:main-log:${applicationLog}`);
    await app.evaluate(({ app: electronApp }) => electronApp.exit(0)).catch(() => undefined);
    await app.close();
    await rm(userData, { recursive: true, force: true });
  }
});

test('desktop-app-lifecycle / isolated tray harness / opens a visible window only when explicitly enabled', async () => {
  const root = path.resolve(process.cwd());
  const userData = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-tray-e2e-'));
  const app = await electron.launch({
    args: [root],
    cwd: root,
    env: {
      ...process.env,
      REPORTER_USER_DATA_DIR: userData,
      REPORTER_TEST_TRAY: '1',
    },
  });

  try {
    const window = await app.firstWindow({ timeout: 8_000 });
    const nativeWindow = await app.browserWindow(window);
    await expect.poll(() => nativeWindow.evaluate((candidate) => candidate.isVisible())).toBe(true);
    await expect(window.getByRole('heading', { name: '股市記者小幫手' })).toBeVisible();
    await expect.poll(async () => {
      const log = await readFile(path.join(userData, 'logs', 'application.log'), 'utf8');
      return (log.match(/"event":"main-ready"/g) ?? []).length;
    }).toBe(1);
    await expect.poll(async () => {
      const database = await access(path.join(userData, 'reporter.sqlite3')).then(() => true, () => false);
      return database;
    }).toBe(true);
    await nativeWindow.evaluate((candidate) => candidate.close());
    await expect.poll(() => nativeWindow.evaluate((candidate) => candidate.isVisible())).toBe(false);
    await expect.poll(() => app.evaluate(({ app: electronApp }) => electronApp.isReady())).toBe(true);

    const electronExecutable = require('electron') as string;
    const secondInstance = spawn(electronExecutable, [root], {
      cwd: root,
      windowsHide: true,
      stdio: 'ignore',
      env: { ...process.env, REPORTER_USER_DATA_DIR: userData },
    });
    const secondExitCode = await new Promise<number | null>((resolve, reject) => {
      secondInstance.once('error', reject);
      secondInstance.once('exit', (code) => resolve(code));
    });
    expect(secondExitCode).toBe(0);
    await expect.poll(() => nativeWindow.evaluate((candidate) => candidate.isVisible())).toBe(true);
    const log = await readFile(path.join(userData, 'logs', 'application.log'), 'utf8');
    expect(log.match(/"event":"main-ready"/g)).toHaveLength(1);
  } finally {
    await app.evaluate(({ app: electronApp }) => electronApp.exit(0)).catch(() => undefined);
    await app.close();
    await rm(userData, { recursive: true, force: true });
  }
});

test('monitoring-schedule / opt-in live-source harness / exposes manual checks without an automatic startup run', async () => {
  const root = path.resolve(process.cwd());
  const userData = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-live-source-e2e-'));
  const app = await electron.launch({
    args: [root], cwd: root,
    env: { ...process.env, REPORTER_USER_DATA_DIR: userData, REPORTER_TEST_LIVE_SOURCES: '1', REPORTER_TEST_TRAY: '1' },
  });
  try {
    const window = await app.firstWindow({ timeout: 8_000 });
    await window.getByRole('tab', { name: '設定' }).click();
    await expect(window.getByRole('button', { name: '立即檢查重大訊息' })).toBeEnabled();
    await expect(window.getByRole('button', { name: '立即檢查違約交割' })).toBeEnabled();
    await expect.poll(async () => {
      const status = await window.evaluate(() => window.reporterApi.getScheduleStatus()) as { available?: boolean };
      return status.available;
    }).toBe(true);
    const Database = require('better-sqlite3') as typeof import('better-sqlite3');
    const database = new Database(path.join(userData, 'reporter.sqlite3'), { readonly: true });
    try { expect(database.prepare('SELECT COUNT(*) AS count FROM job_runs').get()).toEqual({ count: 0 }); }
    finally { database.close(); }
  } finally {
    await app.evaluate(({ app: electronApp }) => electronApp.exit(0)).catch(() => undefined);
    await app.close();
    await rm(userData, { recursive: true, force: true });
  }
});

test('default-disclosure-monitoring / Electron manual check / shows stale official-shaped TPEX data without a false zero success', async () => {
  const root = path.resolve(process.cwd());
  const userData = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-disclosure-e2e-'));
  const twse = await readFile(path.join(root, 'test/fixtures/providers/twse-bfigtu-dashboard-stale.json'), 'utf8');
  const tpex = await readFile(path.join(root, 'test/fixtures/providers/tpex-breach-official-empty.json'), 'utf8');
  const twseReconciliation = await readFile(path.join(root, 'test/fixtures/providers/twse-material-reconciliation.json'), 'utf8');
  const tpexReconciliation = await readFile(path.join(root, 'test/fixtures/providers/tpex-material-reconciliation.json'), 'utf8');
  const requestedTargets: string[] = [];
  const server = createServer((request, response) => {
    const target = new URL(request.url ?? '/', 'http://127.0.0.1').searchParams.get('target') ?? '';
    requestedTargets.push(target);
    response.setHeader('Content-Type', 'application/json; charset=utf-8');
    if (target.includes('/announcement/BFIGTU')) response.end(twse);
    else if (target === 'https://www.tpex.org.tw/www/zh-tw/bulletin/breach') response.end(tpex);
    else if (target.includes('/opendata/t187ap04_L')) response.end(twseReconciliation);
    else if (target.includes('/openapi/v1/mopsfin_t187ap04_O')) response.end(tpexReconciliation);
    else { response.statusCode = 404; response.end('{}'); }
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  if (!address || typeof address === 'string') throw new Error('無法啟動離線來源測試伺服器');
  let app: Awaited<ReturnType<typeof electron.launch>> | undefined;
  try {
    app = await electron.launch({
      args: [root], cwd: root,
      env: {
        ...process.env,
        REPORTER_USER_DATA_DIR: userData,
        REPORTER_TEST_TRAY: '1',
        REPORTER_TEST_LIVE_SOURCES: '1',
        REPORTER_TEST_HTTP_PROXY: `http://127.0.0.1:${address.port}/source`,
      },
    });
    const window = await app.firstWindow({ timeout: 8_000 });
    await window.getByRole('tab', { name: '設定' }).click();
    await window.getByRole('button', { name: '立即檢查違約交割' }).click();
    await expect.poll(async () => {
      const status = await window.evaluate(() => window.reporterApi.getDisclosureMonitorStatus()) as { status: string } | null;
      return status?.status;
    }).toBe('stale');
    await expect(window.getByText('違約交割官方資料尚未更新。')).toBeVisible();
    expect(requestedTargets).toContain('https://www.tpex.org.tw/www/zh-tw/bulletin/breach');
    expect(requestedTargets.some((target) => target.startsWith('https://www.twse.com.tw/rwd/zh/announcement/BFIGTU?startDate=')
      && target.includes('&endDate=') && target.endsWith('&response=json'))).toBe(true);
    expect(requestedTargets).toHaveLength(4);
    await window.getByRole('tab', { name: '違約交割' }).click();
    await expect(window.locator('[aria-labelledby="disclosure-title"] .section-heading .count')).toHaveText('未確認');
    const Database = require('better-sqlite3') as typeof import('better-sqlite3');
    const database = new Database(path.join(userData, 'reporter.sqlite3'), { readonly: true });
    try {
      expect(database.prepare("SELECT status, data_date AS dataDate FROM source_checks WHERE source = 'TPEX'").get())
        .toEqual({ status: 'stale', dataDate: '2026-09-24' });
      expect(database.prepare("SELECT status, data_date AS dataDate FROM source_checks WHERE source = 'TWSE'").get())
        .toEqual({ status: 'stale', dataDate: '2026-09-24' });
    } finally { database.close(); }
  } finally {
    if (app) {
      await app.evaluate(({ app: electronApp }) => electronApp.exit(0)).catch(() => undefined);
      await app.close();
    }
    await new Promise<void>((resolve) => server.close(() => resolve()));
    await rm(userData, { recursive: true, force: true });
  }
});
