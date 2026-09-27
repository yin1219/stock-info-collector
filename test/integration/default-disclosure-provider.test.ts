import { describe, expect, it } from 'vitest';
import { createDefaultDisclosureProvider } from '../../src/providers/default-disclosures';
import { FakeHttpClient, loadFixture } from '../support';

describe('default-disclosure-monitoring / provider contract', () => {
  it('reads the TWSE dashboard BFIGTU range and dates an empty latest report from its summary table', async () => {
    const requests: string[] = [];
    const officialShape = {
      stat: 'OK',
      date: '20260801', // Query date is not the latest reporting date.
      tables: [
        { fields: ['申報日期', '買進、賣出合計總金額'], data: [['115/09/23', '1,000'], ['115/09/24', '2,000']] },
        { fields: ['申報日期', '證券代號', '證券名稱', '證券商名稱', '個股違約總金額(註1)'], data: [
          ['115.09.21', '6213', '舊日公司', '測試券商', '319,306,000'],
          ['總計', '', '', '', '319,306,000'],
        ] },
      ],
    };
    const provider = createDefaultDisclosureProvider({ async get(url) { requests.push(url); return officialShape; } }, 'TWSE');
    await expect(provider.fetchForDate('2026-09-27')).resolves.toMatchObject({ market: 'TWSE', dataDate: '2026-09-24', records: [] });
    expect(requests).toEqual([
      'https://www.twse.com.tw/rwd/zh/announcement/BFIGTU?startDate=20260801&endDate=20260927&response=json',
    ]);
  });

  it('keeps only dated TWSE dashboard security rows and fails closed when the summary date is absent', async () => {
    const endpoint = 'fixture://twse/dashboard';
    const fixture = { stat: 'OK', tables: [
      { fields: ['申報日期'], data: [['115/09/26']] },
      { fields: ['申報日期', '證券代號', '證券名稱', '證券商名稱', '個股違約總金額(註1)'], data: [
        ['115.09.26', '2330', '測試公司', '測試券商', '10,000,000'],
        ['115.09.25', '2317', '舊日公司', '測試券商', '20,000,000'],
      ] },
    ] };
    const provider = createDefaultDisclosureProvider(new FakeHttpClient({ [endpoint]: fixture }), 'TWSE', endpoint);
    await expect(provider.fetchForDate('2026-09-26')).resolves.toMatchObject({ dataDate: '2026-09-26', records: [
      { stockCode: '2330', companyName: '測試公司', brokerCode: '測試券商' },
    ] });
    fixture.tables[0].data = [];
    await expect(provider.fetchForDate('2026-09-26')).rejects.toThrow(/申報日期/);
    const malformed = { stat: 'OK', tables: [
      { fields: ['申報日期'], data: [['115/09/26']] },
      { fields: ['申報日期', '證券代號'], data: [null] },
    ] };
    await expect(createDefaultDisclosureProvider(new FakeHttpClient({ [endpoint]: malformed }), 'TWSE', endpoint)
      .fetchForDate('2026-09-26')).rejects.toThrow(/個股資料列格式無效/);
  });

  it('uses the TPEX bulletin breach API and treats its official prior-day empty table as stale', async () => {
    const fixture = JSON.parse(await loadFixture('providers/tpex-breach-official-empty.json'));
    const endpoint = 'https://www.tpex.org.tw/www/zh-tw/bulletin/breach';
    const requests: string[] = [];
    const http = { async get(url: string) { requests.push(url); return fixture; } };
    const result = await createDefaultDisclosureProvider(http, 'TPEX').fetchForDate('2026-09-26');
    expect(requests).toEqual([endpoint]);
    expect(result).toMatchObject({ market: 'TPEX', dataDate: '2026-09-24', records: [] });
  });

  it('maps TPEX bulletin breach security rows and rejects a missing security table', async () => {
    const endpoint = 'fixture://tpex/breach';
    const officialShape = JSON.parse(await loadFixture('providers/tpex-breach-official-empty.json'));
    officialShape.date = '20260926~20260926';
    officialShape.tables[1].totalCount = 1;
    officialShape.tables[1].data = [['115/09/26', '測試晶圓', '6488', '測試證券商', '12,000,000']];
    const result = await createDefaultDisclosureProvider(new FakeHttpClient({ [endpoint]: officialShape }), 'TPEX', endpoint)
      .fetchForDate('2026-09-26');
    expect(result).toMatchObject({ dataDate: '2026-09-26', records: [
      { market: 'TPEX', disclosureDate: '2026-09-26', stockCode: '6488', companyName: '測試晶圓', brokerCode: '測試證券商' },
    ] });
    expect(result.records[0].content).toMatchObject({ '個股違約總金額(註1)': '12,000,000' });
    officialShape.tables.pop();
    await expect(createDefaultDisclosureProvider(new FakeHttpClient({ [endpoint]: officialShape }), 'TPEX', endpoint)
      .fetchForDate('2026-09-26')).rejects.toThrow(/個股.*表格/);
  });

  it('normalizes all TWSE BFIGTU records and preserves the response date', async () => {
    const fixture = JSON.parse(await loadFixture('providers/twse-default-disclosures.json'));
    const provider = createDefaultDisclosureProvider(new FakeHttpClient({ 'fixture://twse/BFIGTU': fixture }), 'TWSE', 'fixture://twse/BFIGTU');
    const result = await provider.fetchForDate('2026-09-26');
    expect(result).toMatchObject({ dataDate: '2026-09-26', records: [
      { market: 'TWSE', stockCode: '2330', brokerCode: 'B001' },
      { market: 'TWSE', stockCode: '2317', brokerCode: 'B002' },
    ] });
  });

  it('does not call the live TWSE no-match envelope a dated zero-result', async () => {
    const provider = createDefaultDisclosureProvider(new FakeHttpClient({
      'fixture://twse/no-match': { stat: '很抱歉，沒有符合條件的資料!' },
    }), 'TWSE', 'fixture://twse/no-match');
    await expect(provider.fetchForDate('2026-09-27'))
      .rejects.toThrow(/TWSE.*未提供資料日期.*不能視為本日零筆/);
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
