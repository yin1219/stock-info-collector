import { afterEach, describe, expect, it } from 'vitest';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { addListedCompany } from '../../src/services/watchlist';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
afterEach(async () => {
  await paths?.cleanup();
  paths = undefined;
});

describe('watchlist-management / add / saves a valid listed company', () => {
  it('persists an idempotent company and active watchlist entry in SQLite', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const directory = { async findByStockCode(market: 'TWSE' | 'TPEX', stockCode: string) {
      return stockCode === '2330' ? { market, stockCode, name: '台積電' } : undefined;
    } };
    const service = addListedCompany({
      directory,
      repository: {
        transaction: repositories.transaction,
        saveCompany(company) {
          const { id: _generatedId, ...data } = company;
          return repositories.companies.upsert(data).id;
        },
        saveWatchlist: repositories.watchlist.upsert,
      },
      now: () => '2026-09-26T00:00:00.000Z',
    });

    const first = await service({ market: 'TWSE', stockCode: '2330' });
    const repeated = await service({ market: 'TWSE', stockCode: '2330' });

    expect(repeated.companyId).toBe(first.companyId);
    expect(database.prepare('SELECT COUNT(*) AS count FROM companies').get()).toEqual({ count: 1 });
    expect(database.prepare('SELECT COUNT(*) AS count FROM watchlist_entries').get()).toEqual({ count: 1 });
    expect(repositories.watchlist.find(first.companyId)).toMatchObject({ active: true });
    database.close();
  });

  it('rolls back company creation if its watchlist entry cannot be persisted', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const service = addListedCompany({
      directory: { async findByStockCode(market, stockCode) { return { market, stockCode, name: '台積電' }; } },
      repository: {
        transaction: repositories.transaction,
        saveCompany(company) {
          const { id: _generatedId, ...data } = company;
          return repositories.companies.upsert(data).id;
        },
        saveWatchlist() { throw new Error('simulated persistence failure'); },
      },
    });

    await expect(service({ market: 'TWSE', stockCode: '2330' })).rejects.toThrow('simulated persistence failure');
    expect(database.prepare('SELECT COUNT(*) AS count FROM companies').get()).toEqual({ count: 0 });
    database.close();
  });
});
