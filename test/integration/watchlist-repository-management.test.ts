import { afterEach, describe, expect, it } from 'vitest';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
let database: ReturnType<typeof openDatabase> | undefined;
afterEach(async () => {
  database?.close();
  database = undefined;
  await paths?.cleanup();
  paths = undefined;
});

describe('watchlist-management / deactivate / retains historical events when disabling a company', () => {
  it('updates active state and details, and removes only the watchlist membership', async () => {
    paths = await createTestPaths();
    database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '台積電', updatedAt: '2026-09-26T00:00:00Z' });
    const entry = repositories.watchlist.upsert({
      companyId: company.id,
      active: true,
      category: '',
      notes: '',
      createdAt: '2026-09-26T00:00:00Z',
      updatedAt: '2026-09-26T00:00:00Z',
    });
    const event = repositories.materialEvents.upsert({
      companyId: company.id,
      sourceKey: 'mops:TWSE:2330:20260926:1',
      contentFingerprint: 'hash',
      title: '董事會決議',
      content: '公告全文',
      publishedAt: '2026-09-26T00:00:00Z',
    });

    expect(repositories.watchlist.list()).toContainEqual(expect.objectContaining({
      companyId: company.id, market: 'TWSE', stockCode: '2330', active: true,
    }));
    repositories.watchlist.updateDetails(company.id, { category: '半導體', notes: '重點公司' }, '2026-09-26T00:01:00Z');
    repositories.watchlist.remove(company.id, '2026-09-26T00:02:00Z');

    expect(repositories.watchlist.list({ activeOnly: true })).toEqual([]);
    expect(repositories.watchlist.find(company.id)).toMatchObject({ active: false, category: '半導體', notes: '重點公司' });
    expect(repositories.watchlist.isWatched('TWSE', '2330')).toBe(true);
    expect(repositories.materialEvents.find(event.id)).toMatchObject({ title: '董事會決議', content: '公告全文' });
    expect(entry.companyId).toBe(company.id);
  });
});
