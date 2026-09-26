import { describe, expect, it } from 'vitest';
import { addListedCompany, removeWatchlistCompany, setWatchlistCompanyActive, updateWatchlistCompany, type CompanyDirectory, type WatchlistRepositoryPort } from '../../src/services/watchlist';

describe('watchlist-management / add / saves a valid listed company', () => {
  it('looks up a listed stock code and stores company plus watchlist entry atomically', async () => {
    const writes: string[] = [];
    const service = addListedCompany({
      directory: {
        async findByStockCode(market, stockCode) {
          return { market, stockCode, name: '台積電' };
        },
      },
      repository: {
        transaction(work) { writes.push('begin'); const result = work(); writes.push('commit'); return result; },
        saveCompany(company) { writes.push(`company:${company.stockCode}`); return 'company-1'; },
        saveWatchlist(entry) { writes.push(`watchlist:${entry.companyId}`); return entry; },
      },
      now: () => '2026-09-26T00:00:00.000Z',
    });

    await expect(service({ market: 'TWSE', stockCode: ' 2330 ' })).resolves.toMatchObject({
      companyId: 'company-1',
      stockCode: '2330',
      name: '台積電',
      active: true,
    });
    expect(writes).toEqual(['begin', 'company:2330', 'watchlist:company-1', 'commit']);
  });
});

describe('watchlist-management / deactivate / retains historical events when disabling a company', () => {
  it('edits metadata and removes membership by deactivating instead of deleting', () => {
    const calls: unknown[][] = [];
    const repository = {
      updateDetails(companyId: string, details: { category: string; notes: string }, updatedAt: string) {
        calls.push(['details', companyId, details, updatedAt]);
        return { companyId, ...details };
      },
      remove(companyId: string, updatedAt: string) {
        calls.push(['deactivate', companyId, updatedAt]);
        return { companyId, active: false };
      },
      setActive(companyId: string, active: boolean, updatedAt: string) {
        calls.push(['active', companyId, active, updatedAt]);
        return { companyId, active };
      },
    };

    expect(updateWatchlistCompany(repository, 'company-1', { category: ' 半導體 ', notes: '追蹤' }, '2026-09-26T00:00:00Z'))
      .toEqual({ companyId: 'company-1', category: '半導體', notes: '追蹤' });
    expect(removeWatchlistCompany(repository, 'company-1', '2026-09-26T00:01:00Z')).toEqual({ companyId: 'company-1', active: false });
    expect(setWatchlistCompanyActive(repository, 'company-1', false, '2026-09-26T00:02:00Z')).toEqual({ companyId: 'company-1', active: false });
    expect(calls).toEqual([
      ['details', 'company-1', { category: '半導體', notes: '追蹤' }, '2026-09-26T00:00:00Z'],
      ['deactivate', 'company-1', '2026-09-26T00:01:00Z'],
      ['active', 'company-1', false, '2026-09-26T00:02:00Z'],
    ]);
  });
});

describe('watchlist-management / validation / rejects an unknown stock code', () => {
  it('does not persist an unknown but well-formed stock code', async () => {
    let persisted = false;
    const directory: CompanyDirectory = { async findByStockCode() { return undefined; } };
    const repository: WatchlistRepositoryPort = {
      transaction(work) { return work(); },
      saveCompany() { persisted = true; return 'unexpected'; },
      saveWatchlist() { persisted = true; return {} as never; },
    };

    await expect(addListedCompany({ directory, repository })({ market: 'TWSE', stockCode: '9999' }))
      .rejects.toThrow('目前公司名錄找不到股票代號 9999');
    expect(persisted).toBe(false);
  });

  it('rejects malformed codes and directory results for another market without persistence', async () => {
    let writes = 0;
    const repository: WatchlistRepositoryPort = {
      transaction(work) { return work(); },
      saveCompany() { writes += 1; return 'company'; },
      saveWatchlist() { writes += 1; return {}; },
    };
    await expect(addListedCompany({ directory: { async findByStockCode() { return undefined; } }, repository })({ market: 'TWSE', stockCode: '23A0' }))
      .rejects.toThrow('4 至 6 位數字');
    await expect(addListedCompany({ directory: { async findByStockCode() { return { market: 'TPEX', stockCode: '2330', name: '錯誤市場' }; } }, repository })({ market: 'TWSE', stockCode: '2330' }))
      .rejects.toThrow('市場或股票代號不符');
    expect(writes).toBe(0);
  });
});
