import { describe, expect, it } from 'vitest';
import { createCompanyRegistryProvider, parseCompanyRegistry } from '../../src/providers/company-registry';
import { FakeHttpClient, loadFixture } from '../support';

describe('watchlist-management / provider contract / resolves current TWSE and TPEX company listings', () => {
  it('normalizes exchange-specific official fields and keeps market identity independent', async () => {
    const twse = JSON.parse(await loadFixture('providers/twse-companies.json')) as unknown[];
    const tpex = JSON.parse(await loadFixture('providers/tpex-companies.json')) as unknown[];

    expect(parseCompanyRegistry(twse, 'TWSE')).toContainEqual({
      market: 'TWSE', stockCode: '2330', name: '台灣積體電路製造股份有限公司',
    });
    expect(parseCompanyRegistry(tpex, 'TPEX')).toContainEqual({
      market: 'TPEX', stockCode: '6488', name: '環球晶圓股份有限公司',
    });
    expect(parseCompanyRegistry(twse, 'TWSE').some(({ market, stockCode }) => market === 'TPEX' && stockCode === '6488')).toBe(false);
  });
});

describe('watchlist-management / provider contract / validates by market and rejects an unlisted code', () => {
  it('does not treat a same-number listing in another market as a match', async () => {
    const twse = JSON.parse(await loadFixture('providers/twse-companies.json')) as unknown[];
    const tpex = JSON.parse(await loadFixture('providers/tpex-companies.json')) as unknown[];
    const http = new FakeHttpClient({ 'fixture://twse/company': twse, 'fixture://tpex/company': tpex });
    const provider = createCompanyRegistryProvider(http, {
      TWSE: 'fixture://twse/company',
      TPEX: 'fixture://tpex/company',
    });

    await expect(provider.findByStockCode('TWSE', '6488')).resolves.toMatchObject({ market: 'TWSE', name: '錯誤市場測試股份有限公司' });
    await expect(provider.findByStockCode('TPEX', '6488')).resolves.toMatchObject({ market: 'TPEX', name: '環球晶圓股份有限公司' });
    await expect(provider.findByStockCode('TPEX', '9999')).resolves.toBeUndefined();
    expect(http.requests).toEqual(['fixture://twse/company', 'fixture://tpex/company']);
  });
});

describe('watchlist-management / provider contract / rejects invalid registry payloads', () => {
  it('fails closed for an unavailable or malformed market registry', async () => {
    expect(() => parseCompanyRegistry({ message: 'service unavailable' }, 'TWSE')).toThrow(/公司名錄/);
    const provider = createCompanyRegistryProvider(new FakeHttpClient({ 'fixture://twse/company': new Error('offline') }), {
      TWSE: 'fixture://twse/company', TPEX: 'fixture://tpex/company',
    });
    await expect(provider.findByStockCode('TWSE', '2330')).rejects.toThrow('offline');
  });
});
