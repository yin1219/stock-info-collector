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
  const longTitle = '公告本公司名稱由「舊公司名稱」更名為「新的完整公司名稱」，並補充說明董事會決議、變更登記及其他相關事項';
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
      addEvent.run('event-other', 'company-other', 'other-1', 'hash-2', longTitle, '其他公告全文', '2026-09-27T02:00:00Z');
    } finally { database.close(); }
    await window.getByRole('tab', { name: '重大訊息' }).click();
    await expect(window.getByText('關注 1 筆')).toBeVisible();
    await expect(window.getByRole('button', { name: '台積電 關注公告' })).toBeVisible();
    await expect(window.getByRole('button', { name: `其他公司 ${longTitle}` })).toHaveCount(0);
    await window.getByRole('checkbox', { name: '顯示全部公告' }).check();
    await expect(window.getByText('全部 2 筆')).toBeVisible();
    const otherRow = window.getByRole('button', { name: `其他公司 ${longTitle}` }).locator('..');
    await otherRow.getByRole('button', { name: `其他公司 ${longTitle}` }).click();
    await expect(otherRow.getByRole('complementary', { name: '重大訊息完整內容' })).toContainText('其他公告全文');
    const close = otherRow.getByRole('button', { name: '關閉重大訊息詳情' });
    const closeBounds = await close.boundingBox();
    expect(closeBounds).not.toBeNull();
    expect(closeBounds!.width).toBeGreaterThanOrEqual(48);
    expect(closeBounds!.height).toBeLessThanOrEqual(48);
    const watchedRow = window.getByRole('button', { name: '台積電 關注公告' }).locator('..');
    await watchedRow.getByRole('button', { name: '台積電 關注公告' }).click();
    await expect(watchedRow.getByRole('complementary', { name: '重大訊息完整內容' })).toContainText('關注公告全文');
    await expect(otherRow.getByRole('complementary', { name: '重大訊息完整內容' })).toHaveCount(0);
    window.once('dialog', (dialog) => dialog.accept());
    await watchedRow.getByRole('button', { name: '刪除此筆本機公告' }).click();
    await expect(window.getByRole('button', { name: '台積電 關注公告' })).toHaveCount(0);
    await expect(window.getByRole('button', { name: `其他公司 ${longTitle}` })).toBeVisible();
    const saved = new Database(path.join(userData, 'reporter.sqlite3'), { readonly: true });
    try {
      expect(saved.prepare('SELECT deleted_at AS deletedAt FROM material_events WHERE id = ?').get('event-watched'))
        .toMatchObject({ deletedAt: expect.any(String) });
      expect(saved.prepare('SELECT COUNT(*) AS count FROM material_event_deletion_audit WHERE event_id = ?').get('event-watched'))
        .toEqual({ count: 1 });
      expect(saved.prepare('SELECT deleted_at AS deletedAt FROM material_events WHERE id = ?').get('event-other'))
        .toEqual({ deletedAt: null });
    } finally { saved.close(); }
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
    await window.getByRole('button', { name: '發送測試通知', exact: true }).click();
    await expect(window.getByText('隔離測試模式不顯示 Windows 通知。')).toBeVisible();
    await window.getByRole('button', { name: '一分鐘後發送測試通知' }).click();
    await expect(window.getByText('隔離測試模式不顯示 Windows 通知。')).toBeVisible();
    await window.getByRole('tab', { name: '重大訊息' }).click();
    await app.evaluate(({ BrowserWindow }) => {
      BrowserWindow.getAllWindows()[0].webContents.send('notification:navigate', { type: 'test-notification' });
    });
    await expect(window.getByRole('heading', { name: '監控排程' })).toBeVisible();
    await expect(window.getByText('已從測試通知返回設定。')).toBeVisible();
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

test('monitoring-schedule / Electron settings / keeps time and frequency hit targets inside the compact panel', async () => {
  const root = path.resolve(process.cwd());
  const userData = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-settings-e2e-'));
  const app = await electron.launch({
    args: [root], cwd: root,
    env: { ...process.env, REPORTER_USER_DATA_DIR: userData, REPORTER_TEST_TRAY: '1' },
  });
  try {
    const window = await app.firstWindow({ timeout: 8_000 });
    await window.setViewportSize({ width: 650, height: 700 });
    await window.getByRole('tab', { name: '設定' }).click();
    for (const name of ['重大訊息開始時間', '重大訊息結束時間', '重大訊息檢查頻率', '違約交割檢查時間']) {
      const control = window.getByLabel(name);
      await control.scrollIntoViewIfNeeded();
      const layout = await control.evaluate((element) => {
        const rect = element.getBoundingClientRect();
        const label = element.closest('label')!.getBoundingClientRect();
        const hit = document.elementFromPoint(rect.right - 12, rect.top + rect.height / 2);
        return {
          fillsLabel: rect.width >= label.width - 1,
          withinLabel: rect.right <= label.right + 1,
          hittable: hit === element || element.contains(hit),
        };
      });
      expect(layout, name).toEqual({ fillsLabel: true, withinLabel: true, hittable: true });
      await control.click({ position: { x: (await control.boundingBox())!.width - 12, y: 21 } });
      await window.keyboard.press('Escape');
    }
    await window.setViewportSize({ width: 1500, height: 900 });
    const materialGroup = window.getByRole('region', { name: '重大訊息檢查' });
    const disclosureGroup = window.getByRole('region', { name: '違約交割檢查' });
    await expect(materialGroup.getByLabel('重大訊息檢查頻率')).toBeVisible();
    await expect(disclosureGroup.getByLabel('違約交割檢查時間')).toBeVisible();
    const frequencyBox = await materialGroup.getByLabel('重大訊息檢查頻率').boundingBox();
    expect(frequencyBox!.width).toBeLessThanOrEqual(220);
    const frequencyArrow = await materialGroup.getByLabel('重大訊息檢查頻率').evaluate((select) => {
      const field = select.closest('label')!;
      const arrow = getComputedStyle(field, '::after');
      const input = getComputedStyle(select);
      return {
        customArrow: input.appearance === 'none' && arrow.content !== 'none',
        rightGap: Number.parseFloat(arrow.right),
        pointerEvents: arrow.pointerEvents,
        textClearance: Number.parseFloat(input.paddingRight),
      };
    });
    expect(frequencyArrow.customArrow).toBe(true);
    expect(frequencyArrow.rightGap).toBeGreaterThanOrEqual(14);
    expect(frequencyArrow.pointerEvents).toBe('none');
    expect(frequencyArrow.textClearance).toBeGreaterThanOrEqual(36);
    await materialGroup.scrollIntoViewIfNeeded();
    await window.screenshot({ path: path.join(root, 'artifacts', 'settings-layout-e2e.png') });
  } finally {
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
  const requests: string[] = [];
  const server = createServer((request, response) => {
    requests.push(request.url ?? '');
    response.statusCode = 503;
    response.end('Offline test fixture: unexpected source request');
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  if (!address || typeof address === 'string') throw new Error('無法啟動離線來源測試伺服器');
  let app: Awaited<ReturnType<typeof electron.launch>> | undefined;
  try {
    app = await electron.launch({
      args: [root], cwd: root,
      env: { ...process.env, REPORTER_USER_DATA_DIR: userData, REPORTER_TEST_LIVE_SOURCES: '1',
        REPORTER_TEST_TRAY: '1', REPORTER_TEST_HTTP_PROXY: `http://127.0.0.1:${address.port}/source` },
    });
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
    expect(requests).toEqual([]);
  } finally {
    if (app) {
      await app.evaluate(({ app: electronApp }) => electronApp.exit(0)).catch(() => undefined);
      await app.close();
    }
    await new Promise<void>((resolve) => server.close(() => resolve()));
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
