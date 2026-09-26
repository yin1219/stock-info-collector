import { describe, expect, it } from 'vitest';
import { createMopsMaterialProvider, parseMopsMaterialHtml } from '../../src/providers/mops-material';
import { FakeHttpClient, loadFixture } from '../support';

describe('material-event-monitoring / provider contract / parses official daily material events', () => {
  it('normalizes market, company, publication time, subject, stable identity, and source link', async () => {
    const html = await loadFixture('providers/mops-material.html');

    expect(parseMopsMaterialHtml(html)).toEqual([
      {
        source: 'mops',
        market: 'TWSE',
        stockCode: '2330',
        companyName: '測試半導體',
        publishedAt: '2026-09-26T09:30:00.000Z',
        title: '董事會決議',
        content: '董事會決議',
        sourceKey: 'mops:TWSE:2330:1150926:173000:2',
        sourceUrl: 'https://mops.twse.com.tw/mops/web/ajax_t05st01?stock=2330&date=1150926&time=173000&seq=2',
        revisionOf: null,
      },
      {
        source: 'mops',
        market: 'TPEX',
        stockCode: '6488',
        companyName: '測試晶圓',
        publishedAt: '2026-09-26T09:45:00.000Z',
        title: '公告營運事項',
        content: '公告營運事項',
        sourceKey: 'mops:TPEX:6488:1150926:174500:1',
        sourceUrl: 'https://mops.twse.com.tw/mops/web/ajax_t05st01?stock=6488&date=1150926&time=174500&seq=1',
        revisionOf: null,
      },
    ]);
  });

  it('rejects a structurally invalid response instead of reporting a false empty success', () => {
    expect(() => parseMopsMaterialHtml('<html><body><p>維護中</p></body></html>')).toThrow(/重大訊息表格/);
  });
});

describe('material-event-monitoring / provider contract / uses a fake HTTP transport', () => {
  it('formats a ROC date request and consumes only an injected fixture response', async () => {
    const html = await loadFixture('providers/mops-material.html');
    const detail = await loadFixture('providers/mops-material-detail.html');
    const endpoint = 'fixture://mops/ajax_t05st02';
    const http = new FakeHttpClient({
      [`${endpoint}?year=115&month=09&day=26`]: html,
      'fixture://mops/mops/web/ajax_t05st01?stock=2330&date=1150926&time=173000&seq=2': detail,
      'fixture://mops/mops/web/ajax_t05st01?stock=6488&date=1150926&time=174500&seq=1': detail,
    });
    const provider = createMopsMaterialProvider(http, endpoint);

    const result = await provider.fetchForDate('2026-09-26');

    expect(http.requests).toEqual([
      `${endpoint}?year=115&month=09&day=26`,
      'fixture://mops/mops/web/ajax_t05st01?stock=2330&date=1150926&time=173000&seq=2',
      'fixture://mops/mops/web/ajax_t05st01?stock=6488&date=1150926&time=174500&seq=1',
    ]);
    expect(result.events).toHaveLength(2);
    expect(result.events[0].content).toContain('董事會通過測試用營運計畫');
    expect(result.events[0].content).not.toBe(result.events[0].title);
    expect(result.status).toBe('complete');
  });
});
