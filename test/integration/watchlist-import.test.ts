import { describe, expect, it } from 'vitest';
import { applyLegacyWatchlistImport, previewLegacyWatchlistImport } from '../../src/services/watchlist-import';
import { loadFixture } from '../support';

describe('watchlist-management / import / reports added skipped and invalid rows idempotently', () => {
  it('previews outcomes without writes, imports partial successes, and skips them on the next run', async () => {
    const configText = await loadFixture('watchlist/config-default.json');
    const directory = {
      async findByStockCode(market: 'TWSE' | 'TPEX', stockCode: string) {
        if (market === 'TWSE' && stockCode === '2330') return { market, stockCode, name: '測試上市公司' };
        if (market === 'TPEX' && stockCode === '6488') return { market, stockCode, name: '測試上櫃公司' };
        return undefined;
      },
    };
    const watched = new Set<string>();
    const preview = await previewLegacyWatchlistImport(configText, {
      directory,
      isWatched: (market, code) => watched.has(`${market}:${code}`),
    });

    expect(preview.summary).toEqual({ added: 2, skipped: 1, failed: 2 });
    expect(watched.size).toBe(0);

    const repository = {
      transaction<T>(work: () => T): T { return work(); },
      isWatched: (market: string, code: string) => watched.has(`${market}:${code}`),
      saveCompany: (listing: { market: string; stockCode: string; name: string }) => `${listing.market}:${listing.stockCode}`,
      saveWatchlist(entry: { market: string; stockCode: string }) { watched.add(`${entry.market}:${entry.stockCode}`); },
    };
    const first = await applyLegacyWatchlistImport(configText, { directory, repository, now: () => '2026-09-26T00:00:00.000Z' });
    const second = await applyLegacyWatchlistImport(configText, { directory, repository, now: () => '2026-09-26T00:00:00.000Z' });

    expect(first.summary).toEqual({ added: 2, skipped: 1, failed: 2 });
    expect(second.summary).toEqual({ added: 0, skipped: 3, failed: 2 });
    expect(watched).toEqual(new Set(['TWSE:2330', 'TPEX:6488']));
  });

  it('continues after one registry failure and one persistence failure while skipping an existing company', async () => {
    const directory = {
      async findByStockCode(market: 'TWSE' | 'TPEX', stockCode: string) {
        if (stockCode === '2330') {
          if (market === 'TWSE') throw new Error('TWSE temporarily offline');
          return { market, stockCode, name: '上櫃公司' };
        }
        if (stockCode === '2331') throw new Error('both registries offline');
        if (stockCode === '2332' || stockCode === '2333') return { market, stockCode, name: `公司${stockCode}` };
        return undefined;
      },
    };
    const report = await applyLegacyWatchlistImport('{"StockNumbers":"2330,2331,2332,2333"}', {
      directory,
      repository: {
        transaction<T>(work: () => T): T { return work(); },
        isWatched(_market, stockCode) { return stockCode === '2332'; },
        saveCompany(listing) {
          if (listing.stockCode === '2333') throw new Error('database write failed');
          return `${listing.market}:${listing.stockCode}`;
        },
        saveWatchlist() { return undefined; },
      },
      now: () => '2026-09-26T00:00:00.000Z',
    });

    expect(report.summary).toEqual({ added: 1, skipped: 1, failed: 2 });
    expect(report.items.map(({ status }) => status)).toEqual(['add', 'failed', 'skip', 'failed']);
    expect(report.items[1].reason).toContain('both registries offline');
    expect(report.items[3].reason).toContain('database write failed');
  });
});
