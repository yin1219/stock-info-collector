import { describe, expect, it } from 'vitest';
import { createDefaultDisclosureProvider } from '../../src/providers/default-disclosures';
import { FakeHttpClient, loadFixture } from '../support';

describe('default-disclosure-monitoring / provider contract', () => {
  it('normalizes all TWSE BFIGTU records and preserves the response date', async () => {
    const fixture = JSON.parse(await loadFixture('providers/twse-default-disclosures.json'));
    const provider = createDefaultDisclosureProvider(new FakeHttpClient({ 'fixture://twse/BFIGTU': fixture }), 'TWSE', 'fixture://twse/BFIGTU');
    const result = await provider.fetchForDate('2026-09-26');
    expect(result).toMatchObject({ dataDate: '2026-09-26', records: [
      { market: 'TWSE', stockCode: '2330', brokerCode: 'B001' },
      { market: 'TWSE', stockCode: '2317', brokerCode: 'B002' },
    ] });
  });

  it('normalizes TPEX stockDetail rows, permits explicit empty success and rejects malformed payloads', async () => {
    const fixture = JSON.parse(await loadFixture('providers/tpex-default-disclosures.json'));
    const provider = createDefaultDisclosureProvider(new FakeHttpClient({
      'fixture://tpex/violation': fixture,
      'fixture://tpex/multiple': { date: '2026-09-26', stockDetail: [
        { date: '2026-09-26', stockCode: '6488', companyName: '測試晶圓', brokerCode: 'B003' },
        { date: '2026-09-26', stockCode: '8299', companyName: '第二上櫃公司', brokerCode: 'B004' },
      ] },
      'fixture://tpex/empty': { date: '2026-09-26', stockDetail: [] },
      'fixture://tpex/old': { date: '2026-09-25', stockDetail: [] },
      'fixture://tpex/bad': { message: 'invalid' },
    }), 'TPEX', 'fixture://tpex/violation');
    await expect(provider.fetchForDate('2026-09-26')).resolves.toMatchObject({ records: [{ market: 'TPEX', stockCode: '6488', brokerCode: 'B003' }] });
    await expect(createDefaultDisclosureProvider(new FakeHttpClient({ 'fixture://tpex/multiple': {
      date: '2026-09-26', stockDetail: [
        { date: '2026-09-26', stockCode: '6488', companyName: '測試晶圓', brokerCode: 'B003' },
        { date: '2026-09-26', stockCode: '8299', companyName: '第二上櫃公司', brokerCode: 'B004' },
      ],
    } }), 'TPEX', 'fixture://tpex/multiple').fetchForDate('2026-09-26')).resolves.toMatchObject({ records: [
      { stockCode: '6488' }, { stockCode: '8299' },
    ] });
    await expect(createDefaultDisclosureProvider(new FakeHttpClient({ 'fixture://tpex/empty': { date: '2026-09-26', stockDetail: [] } }), 'TPEX', 'fixture://tpex/empty').fetchForDate('2026-09-26')).resolves.toMatchObject({ dataDate: '2026-09-26', records: [] });
    await expect(createDefaultDisclosureProvider(new FakeHttpClient({ 'fixture://tpex/bad': { message: 'invalid' } }), 'TPEX', 'fixture://tpex/bad').fetchForDate('2026-09-26')).rejects.toThrow(/格式/);
    await expect(createDefaultDisclosureProvider(new FakeHttpClient({ 'fixture://tpex/old': { date: '2026-09-25', stockDetail: [] } }), 'TPEX', 'fixture://tpex/old').fetchForDate('2026-09-26')).resolves.toMatchObject({ dataDate: '2026-09-25', records: [] });
  });

  it('does not report old target data as an empty current-day result', async () => {
    const provider = createDefaultDisclosureProvider(new FakeHttpClient({ 'fixture://twse/old': { dataDate: '2026-09-25', data: [] } }), 'TWSE', 'fixture://twse/old');
    await expect(provider.fetchForDate('2026-09-26')).resolves.toMatchObject({ dataDate: '2026-09-25', records: [] });
  });

  it('validates TWSE empty, old-date, malformed-row, and explicit per-row dates', async () => {
    const provider = createDefaultDisclosureProvider(new FakeHttpClient({
      'fixture://twse/empty': { dataDate: '2026-09-26', data: [] },
      'fixture://twse/old': { dataDate: '2026-09-25', data: [] },
      'fixture://twse/bad-row': { dataDate: '2026-09-26', data: [{ 股票代號: 'not-code' }] },
      'fixture://twse/per-row-date': { data: [
        { 日期: '2026-09-26', 股票代號: '2330', 公司名稱: '測試公司' },
        { 日期: '2026-09-25', 股票代號: '2317', 公司名稱: '前日資料' },
      ] },
    }), 'TWSE', 'fixture://twse/empty');
    await expect(provider.fetchForDate('2026-09-26')).resolves.toMatchObject({ dataDate: '2026-09-26', records: [] });
    await expect(createDefaultDisclosureProvider(new FakeHttpClient({ 'fixture://twse/old': { dataDate: '2026-09-25', data: [] } }), 'TWSE', 'fixture://twse/old').fetchForDate('2026-09-26'))
      .resolves.toMatchObject({ dataDate: '2026-09-25', records: [] });
    await expect(createDefaultDisclosureProvider(new FakeHttpClient({ 'fixture://twse/bad-row': { dataDate: '2026-09-26', data: [{ 股票代號: 'not-code' }] } }), 'TWSE', 'fixture://twse/bad-row').fetchForDate('2026-09-26'))
      .rejects.toThrow(/有效股票代號/);
    const dated = await createDefaultDisclosureProvider(new FakeHttpClient({ 'fixture://twse/per-row-date': { data: [
      { 日期: '2026-09-26', 股票代號: '2330', 公司名稱: '測試公司' },
      { 日期: '2026-09-25', 股票代號: '2317', 公司名稱: '前日資料' },
    ] } }), 'TWSE', 'fixture://twse/per-row-date').fetchForDate('2026-09-26');
    expect(dated).toMatchObject({ dataDate: '2026-09-26', records: [{ stockCode: '2330' }] });
  });
});
