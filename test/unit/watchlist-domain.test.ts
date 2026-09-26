import { describe, expect, it } from 'vitest';
import { createCompany, setWatchlistActive, updateWatchlistDetails } from '../../src/domain/watchlist';

describe('watchlist-management / domain / validates and transitions watchlist entities', () => {
  it('creates a normalized listed company and an active watchlist entry', () => {
    const company = createCompany({ market: 'TWSE', stockCode: ' 2330 ', name: '台積電' });

    expect(company).toMatchObject({ market: 'TWSE', stockCode: '2330', name: '台積電' });
    expect(company.id).toBeTruthy();
    expect(company.watchlist).toMatchObject({ active: true, category: '', notes: '' });
  });

  it('rejects malformed company data before it can be persisted', () => {
    expect(() => createCompany({ market: 'TWSE', stockCode: '23X0', name: '測試公司' })).toThrow(/股票代號/);
    expect(() => createCompany({ market: 'OTC' as 'TWSE', stockCode: '1234', name: '測試公司' })).toThrow(/市場/);
    expect(() => createCompany({ market: 'TPEX', stockCode: '1234', name: ' ' })).toThrow(/公司名稱/);
  });

  it('updates notes and active state without removing the company identity', () => {
    const original = createCompany({ market: 'TPEX', stockCode: '6488', name: '環球晶' });
    const detailed = updateWatchlistDetails(original, { category: '半導體', notes: '法說會追蹤' });
    const inactive = setWatchlistActive(detailed, false);

    expect(inactive).toMatchObject({ id: original.id, stockCode: '6488', watchlist: {
      active: false,
      category: '半導體',
      notes: '法說會追蹤',
    } });
    expect(original.watchlist.active).toBe(true);
  });
});
