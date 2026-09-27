import { describe, expect, it } from 'vitest';
import { createMaterialReconciliationProvider } from '../../src/providers/material-reconciliation';
import { FakeHttpClient, loadFixture } from '../support';

describe('material-event-monitoring / provider contract / daily reconciliation providers', () => {
  it('maps TWSE records, preserves announcement content and stable source identity', async () => {
    const data = JSON.parse(await loadFixture('providers/twse-material-reconciliation.json'));
    const http = new FakeHttpClient({ 'fixture://twse/material': data });
    const result = await createMaterialReconciliationProvider(http, 'TWSE', 'fixture://twse/material').fetchForDate('2026-09-26');
    expect(result).toMatchObject({ status: 'complete', dataDate: '2026-09-26' });
    expect(result.events[0]).toMatchObject({ market: 'TWSE', stockCode: '2330', title: '董事會決議', content: 'TWSE 對帳完整內容', sourceKey: 'twse:123' });
    expect(http.requests).toEqual(['fixture://twse/material']);
  });

  it('maps TPEX records and rejects a response without an authoritative data date', async () => {
    const data = JSON.parse(await loadFixture('providers/tpex-material-reconciliation.json'));
    const provider = createMaterialReconciliationProvider(new FakeHttpClient({ 'fixture://tpex/material': data }), 'TPEX', 'fixture://tpex/material');
    const result = await provider.fetchForDate('2026-09-26');
    expect(result).toMatchObject({ dataDate: '2026-09-26', events: [{ market: 'TPEX', stockCode: '6488', publishedAt: '2026-09-26T09:45:00.000Z' }] });

    const bad = createMaterialReconciliationProvider(new FakeHttpClient({ 'fixture://tpex/material': [] }), 'TPEX', 'fixture://tpex/material');
    await expect(bad.fetchForDate('2026-09-26')).rejects.toThrow(/資料日期/);
  });

  it('parses the current official OpenAPI column names and compact announcement times', async () => {
    const twse = createMaterialReconciliationProvider(new FakeHttpClient({ 'fixture://twse/current': [{
      出表日期: '1150927', 發言日期: '1150926', 發言時間: '70003', 公司代號: '2330',
      公司名稱: '測試公司', '主旨 ': '重大訊息範例', 說明: '公開內容範例',
    }] }), 'TWSE', 'fixture://twse/current');
    const tpex = createMaterialReconciliationProvider(new FakeHttpClient({ 'fixture://tpex/current': [{
      Date: '1150926', 發言日期: '1150925', 發言時間: '70003',
      SecuritiesCompanyCode: '6488', CompanyName: '測試上櫃公司', 主旨: '上櫃訊息範例', 說明: '公開內容範例',
    }] }), 'TPEX', 'fixture://tpex/current');

    await expect(twse.fetchForDate('2026-09-27')).resolves.toMatchObject({
      dataDate: '2026-09-27',
      events: [{ stockCode: '2330', publishedAt: '2026-09-26T23:00:03.000Z' }],
    });
    await expect(tpex.fetchForDate('2026-09-26')).resolves.toMatchObject({
      dataDate: '2026-09-26',
      events: [{ stockCode: '6488', publishedAt: '2026-09-25T23:00:03.000Z' }],
    });
  });
});
