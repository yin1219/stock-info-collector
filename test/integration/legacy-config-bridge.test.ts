import { afterEach, describe, expect, it } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createLegacyConfigBridge, legacyConfigPathForPlan } from '../../src/main/legacy-config-bridge';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
afterEach(async () => { await paths?.cleanup(); paths = undefined; });

async function setup() {
  paths = await createTestPaths();
  const configPath = path.join(paths.root, 'default.json');
  await writeFile(configPath, '{"Debug":false,"LogLevel":"info","StockNumbers":"2330","Extra":{"keep":true}}\n', 'utf8');
  const database = openDatabase(paths.database);
  const repositories = createRepositories(database);
  const now = '2026-09-29T00:00:00.000Z';
  const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2454', name: '測試公司', updatedAt: now });
  repositories.watchlist.upsert({ companyId: company.id, active: true, category: '', notes: '', createdAt: now, updatedAt: now });
  return { configPath, database, repositories, now };
}

describe('watchlist-management / v1 config bridge / confirmed active codes only', () => {
  it('keeps a recognized but disabled v1 config available for reviewed rollback-list synchronization', async () => {
    const { configPath, database, repositories, now } = await setup();
    try {
      const bridge = createLegacyConfigBridge({
        repositories,
        locateLegacyConfig: async () => legacyConfigPathForPlan({ kind: 'legacy-disabled', configPath }),
        now: () => now,
      });
      const review = await bridge.review();
      expect(review).toMatchObject({ currentCodes: ['2330'], activeCodes: ['2454'], confirmed: false });
      await bridge.confirm(review.hash);
      expect(JSON.parse(await readFile(configPath, 'utf8')).StockNumbers).toBe('2454');
      expect(legacyConfigPathForPlan({ kind: 'v2', legacyConfigPath: configPath })).toBe(configPath);
      expect(legacyConfigPathForPlan({ kind: 'v2' })).toBeNull();
      expect(legacyConfigPathForPlan({ kind: 'blocked', reason: '未知工作' })).toBeNull();
    } finally { database.close(); }
  });

  it('previews both lists without changing v1; confirmation applies and later changes follow', async () => {
    const { configPath, database, repositories, now } = await setup();
    try {
      const bridge = createLegacyConfigBridge({
        repositories, locateLegacyConfig: async () => configPath, now: () => now,
      });
      const review = await bridge.review();
      expect(review).toMatchObject({ currentCodes: ['2330'], activeCodes: ['2454'], confirmed: false });
      expect(JSON.parse(await readFile(configPath, 'utf8')).StockNumbers).toBe('2330');
      expect(await bridge.applyIfConfirmed()).toEqual({ status: 'not-confirmed' });

      const confirmed = await bridge.confirm(review.hash);
      expect(confirmed.status).toBe('applied');
      expect(JSON.parse(await readFile(configPath, 'utf8'))).toMatchObject({
        Debug: false, LogLevel: 'info', StockNumbers: '2454', Extra: { keep: true },
      });
      expect(await readFile(confirmed.backupPath, 'utf8')).toContain('"StockNumbers":"2330"');

      const second = repositories.companies.upsert({ market: 'TPEX', stockCode: '6488', name: '第二間公司', updatedAt: now });
      repositories.watchlist.upsert({ companyId: second.id, active: true, category: '', notes: '', createdAt: now, updatedAt: now });
      expect((await bridge.applyIfConfirmed()).status).toBe('applied');
      expect(JSON.parse(await readFile(configPath, 'utf8')).StockNumbers).toBe('2454,6488');
      repositories.watchlist.setActive(companyId(repositories, '2454'), false, now);
      repositories.watchlist.setActive(companyId(repositories, '6488'), false, now);
      await expect(bridge.applyIfConfirmed()).rejects.toThrow(/至少一間/);
      expect(JSON.parse(await readFile(configPath, 'utf8')).StockNumbers).toBe('2454,6488');
    } finally { database.close(); }
  });

  it('rejects external edits and no longer recognized tasks without claiming a successful apply', async () => {
    const { configPath, database, repositories, now } = await setup();
    let recognized = true;
    try {
      const bridge = createLegacyConfigBridge({
        repositories, locateLegacyConfig: async () => recognized ? configPath : null, now: () => now,
      });
      const review = await bridge.review();
      await bridge.confirm(review.hash);
      await writeFile(configPath, '{"Debug":true,"StockNumbers":"1101"}', 'utf8');
      await expect(bridge.applyIfConfirmed()).rejects.toThrow(/外部修改/);
      expect(JSON.parse(await readFile(configPath, 'utf8')).StockNumbers).toBe('1101');
      recognized = false;
      await expect(bridge.applyIfConfirmed()).rejects.toThrow(/無法辨識/);
    } finally { database.close(); }
  });

  it('requires the same reviewed file and rejects an invalid review hash', async () => {
    const { configPath, database, repositories, now } = await setup();
    try {
      const bridge = createLegacyConfigBridge({ repositories, locateLegacyConfig: async () => configPath, now: () => now });
      await expect(bridge.confirm('stale-review')).rejects.toThrow(/重新核對/);
      expect(JSON.parse(await readFile(configPath, 'utf8')).StockNumbers).toBe('2330');
    } finally { database.close(); }
  });
});

function companyId(repositories: ReturnType<typeof createRepositories>, stockCode: string): string {
  const item = repositories.watchlist.list().find((row) => row.stockCode === stockCode);
  if (!item) throw new Error('fixture company missing');
  return item.companyId;
}
