// Opt-in Playwright canary. This file is intentionally excluded from npm test.
const { _electron: electron } = require('@playwright/test');
const Database = require('better-sqlite3');
const { mkdtemp, mkdir, rm } = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');

const delay = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds));

async function completedJob(databasePath, kind) {
  const deadline = Date.now() + 90_000;
  while (Date.now() < deadline) {
    const database = new Database(databasePath, { readonly: true, fileMustExist: true });
    try {
      const job = database.prepare('SELECT id, status FROM job_runs WHERE kind = ? ORDER BY rowid DESC LIMIT 1').get(kind);
      if (job && job.status !== 'active') {
        const sources = database.prepare(`
          SELECT source, status, data_date AS dataDate, record_count AS recordCount, error_message AS errorMessage
          FROM source_checks WHERE job_run_id = ? ORDER BY source
        `).all(job.id);
        return { status: job.status, sources };
      }
    } finally {
      database.close();
    }
    await delay(250);
  }
  throw new Error(`${kind} 手動檢查未在 90 秒內完成`);
}

(async () => {
  const root = path.resolve(__dirname, '..');
  const userData = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-live-canary-'));
  const output = path.join(root, 'artifacts', 'live-source-canary');
  let app;
  try {
    app = await electron.launch({
      args: [root], cwd: root,
      env: {
        ...process.env,
        REPORTER_USER_DATA_DIR: userData,
        REPORTER_TEST_LIVE_SOURCES: '1',
        REPORTER_TEST_TRAY: '1',
        REPORTER_TEST_HTTP_PROXY: '',
        REPORTER_TEST_LIVE_OAUTH: '',
      },
    });
    const window = await app.firstWindow({ timeout: 15_000 });
    await window.getByRole('tab', { name: '設定' }).click();
    const databasePath = path.join(userData, 'reporter.sqlite3');
    const canaryWatchCode = process.env.REPORTER_CANARY_WATCH_CODE;
    if (canaryWatchCode) {
      if (!/^\d{4,6}$/.test(canaryWatchCode)) throw new Error('REPORTER_CANARY_WATCH_CODE 必須是 4 至 6 位數字');
      const seed = new Database(databasePath);
      try {
        const now = new Date().toISOString();
        seed.prepare('INSERT INTO companies (id, market, stock_code, name, updated_at) VALUES (?, ?, ?, ?, ?)')
          .run(`canary-company-${canaryWatchCode}`, 'TPEX', canaryWatchCode, '驗收關注公司', now);
        seed.prepare('INSERT INTO watchlist_entries (id, company_id, active, category, notes, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)')
          .run(`canary-watch-${canaryWatchCode}`, `canary-company-${canaryWatchCode}`, 1, '', '', now, now);
      } finally {
        seed.close();
      }
    }
    await window.getByRole('button', { name: '立即檢查重大訊息' }).click();
    const material = await completedJob(databasePath, 'material-event-monitoring');
    await window.getByRole('button', { name: '立即檢查違約交割' }).click();
    const disclosure = await completedJob(databasePath, 'default-disclosure-monitoring');
    const reconciliation = await completedJob(databasePath, 'material-event-reconciliation');

    await mkdir(output, { recursive: true });
    await window.getByRole('tab', { name: '重大訊息' }).click();
    const watchedListCount = await window.locator('.material-list .material-row').count();
    await window.getByRole('checkbox', { name: '顯示全部公告' }).check();
    await window.getByText(/全部 \d+ 筆/).waitFor();
    const materialListCount = await window.locator('.material-list .material-row').count();
    const backfillVisible = canaryWatchCode
      ? await window.locator('.material-list .material-row').filter({ hasText: canaryWatchCode }).count()
      : null;
    let materialDetail;
    if (materialListCount > 0) {
      await window.locator('.material-summary').first().click();
      const panel = window.getByRole('complementary', { name: '重大訊息完整內容' });
      await panel.waitFor();
      materialDetail = {
        textLength: (await panel.innerText()).length,
        officialLink: (await panel.getByRole('link', { name: '開啟來源公告' }).getAttribute('href'))?.startsWith('https://mopsov.twse.com.tw/') ?? false,
      };
    }
    const saved = new Database(databasePath, { readonly: true, fileMustExist: true });
    let materialPersistence;
    let backfillSaved = null;
    try {
      materialPersistence = saved.prepare('SELECT COUNT(*) AS count, SUM(CASE WHEN length(content) > length(title) THEN 1 ELSE 0 END) AS fullContentCount FROM material_events').get();
      if (canaryWatchCode) {
        backfillSaved = saved.prepare("SELECT COUNT(*) AS count FROM material_events WHERE company_id = ? AND source = 'tpex-reconciliation'")
          .get(`canary-company-${canaryWatchCode}`).count;
      }
    } finally {
      saved.close();
    }
    await window.screenshot({ path: path.join(output, 'material.png'), fullPage: true });
    await window.getByRole('tab', { name: '違約交割' }).click();
    await window.screenshot({ path: path.join(output, 'disclosures.png'), fullPage: true });
    await window.getByRole('tab', { name: '重大訊息' }).click();
    const afterReconciliationCount = await window.locator('.material-list .material-row').count();
    console.log(JSON.stringify({ material, materialUi: { watchedListCount, listedCount: materialListCount, detail: materialDetail, afterReconciliationCount, backfillVisible }, materialPersistence: { ...materialPersistence, backfillSaved }, disclosure, reconciliation, screenshots: output }, null, 2));
    if (canaryWatchCode && (backfillVisible !== 1 || backfillSaved !== 1)) {
      throw new Error(`對帳補存未同時在 SQLite 與畫面各出現一次：${canaryWatchCode}`);
    }
    if ([material, disclosure, reconciliation].some(({ status }) => status !== 'complete')) {
      console.error('Live-source acceptance incomplete: one or more jobs were not complete.');
      process.exitCode = 1;
    }
  } finally {
    if (app) {
      await app.evaluate(({ app: electronApp }) => electronApp.exit(0)).catch(() => undefined);
      await app.close().catch(() => undefined);
    }
    await rm(userData, { recursive: true, force: true });
  }
})().catch((error) => {
  console.error('Live-source Playwright canary failed:', error.message);
  process.exitCode = 1;
});
