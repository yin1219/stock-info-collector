import { afterEach, describe, expect, it } from 'vitest';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { createWatchlistController } from '../../src/main/watchlist-controller';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
let database: ReturnType<typeof openDatabase> | undefined;
afterEach(async () => {
  database?.close();
  database = undefined;
  await paths?.cleanup();
  paths = undefined;
});

describe('watchlist-management / add / saves a valid listed company', () => {
  it('connects listing validation and watchlist changes to SQLite-backed results immediately', async () => {
    paths = await createTestPaths();
    database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const controller = createWatchlistController({
      repositories,
      directory: {
        async findByStockCode(market, stockCode) {
          if (market === 'TWSE' && stockCode === '2330') return { market, stockCode, name: '台積電' };
          return undefined;
        },
      },
      now: () => '2026-09-26T00:00:00.000Z',
    });

    await expect(controller.add({ market: 'TWSE', stockCode: '2330' })).resolves.toMatchObject({ stockCode: '2330', active: true });
    expect(controller.list({ activeOnly: true })).toMatchObject([{ stockCode: '2330', name: '台積電' }]);
    await expect(controller.add({ market: 'TPEX', stockCode: '2330' })).rejects.toThrow(/找不到股票代號/);
    controller.update({ companyId: controller.list()[0].companyId, category: ' 半導體 ', notes: ' 採訪追蹤 ' });
    controller.setActive({ companyId: controller.list()[0].companyId, active: false });
    expect(controller.list({ activeOnly: true })).toEqual([]);
    expect(controller.list()[0]).toMatchObject({ active: false, category: '半導體', notes: '採訪追蹤' });
  });
});
