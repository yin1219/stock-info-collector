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
});
