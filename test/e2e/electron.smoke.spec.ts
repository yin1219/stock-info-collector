import { access, mkdtemp, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { _electron as electron, expect, test } from '@playwright/test';

test('desktop-app-lifecycle / Electron shell / opens the isolated renderer', async () => {
  const root = path.resolve(process.cwd());
  const userData = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-e2e-'));
  const app = await electron.launch({
    args: [root],
    cwd: root,
    env: {
      ...process.env,
      REPORTER_USER_DATA_DIR: userData,
    },
  });

  try {
    const window = await app.firstWindow({ timeout: 8_000 });
    window.on('pageerror', (error) => console.error(`renderer:pageerror:${error.message}`));
    await expect(window.getByRole('heading', { name: '股市記者小幫手' })).toBeVisible();
    await expect(window).toHaveTitle('股市記者小幫手');
    await expect.poll(() => window.evaluate(() => typeof window.reporterApi.listWatchlist)).toBe('function');
    await expect.poll(() => window.evaluate(() => typeof window.reporterApi.listDisclosures)).toBe('function');
    await expect.poll(() => window.evaluate(() => typeof window.reporterApi.listMaterialEvents)).toBe('function');
    await expect(window.evaluate(() => window.reporterApi.listWatchlist({ activeOnly: true }))).resolves.toEqual([]);
    await expect(window.evaluate(() => window.reporterApi.listDisclosures({ disclosureDate: '2026-09-26' }))).resolves.toEqual([]);
    await expect(window.evaluate(() => window.reporterApi.listMaterialEvents({ unreadOnly: true }))).resolves.toEqual([]);
    await window.getByRole('tab', { name: '違約交割' }).click();
    await expect(window.getByRole('heading', { name: '全市場違約交割' })).toBeVisible();
    await window.getByRole('tab', { name: '重大訊息' }).click();
    await expect(window.getByRole('heading', { name: '重大訊息' })).toBeVisible();
    await access(path.join(userData, 'reporter.sqlite3'));
    const applicationLog = await readFile(path.join(userData, 'logs', 'application.log'), 'utf8');
    expect(applicationLog).toContain('"event":"main-ready"');
    expect(applicationLog).toContain('"event":"renderer-loaded"');
  } finally {
    const applicationLog = await readFile(path.join(userData, 'logs', 'application.log'), 'utf8').catch(() => '');
    if (applicationLog) console.info(`e2e:main-log:${applicationLog}`);
    await app.close();
    await rm(userData, { recursive: true, force: true });
  }
});
