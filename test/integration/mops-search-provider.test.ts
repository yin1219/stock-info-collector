import { describe, expect, it } from 'vitest';
import { createMopsSearchMaterialProvider } from '../../src/providers/mops-search-material';
import { loadFixture } from '../support';

const endpoint = 'https://mopsov.twse.com.tw/mops/web/ezsearch_query';
const detail = 'https://mopsov.twse.com.tw/mops/web/ajax_t05sr01_1?firstin=true&step=1&SEQ_NO=2&SPOKE_TIME=70004&SPOKE_DATE=20260926&COMPANY_ID=2330';

describe('material-event-monitoring / MOPS website search / official dated announcements', () => {
  it('queries both markets by ROC date, gets full detail, and warns that search results can miss daily records', async () => {
    const payload = JSON.parse(await loadFixture('providers/mops-search-material.json')) as { status: string; data: Array<{ TYPEK: string }> };
    const detailHtml = await loadFixture('providers/mops-material-detail.html');
    const requests: Array<{ url: string; body?: string }> = [];
    const provider = createMopsSearchMaterialProvider({
      async post(url, body) {
        requests.push({ url, body });
        const label = new URLSearchParams(body).get('TYPEK') === 'sii' ? '上市' : '上櫃';
        return { ...payload, data: payload.data.filter((row) => row.TYPEK === label) };
      },
      async get(url) { requests.push({ url }); return detailHtml; },
    });

    const result = await provider.fetchForDate('2026-09-26');

    expect(requests.filter((request) => request.body)).toHaveLength(2);
    expect(requests.filter((request) => request.body).map((request) => new URLSearchParams(request.body).get('TYPEK'))).toEqual(['sii', 'otc']);
    for (const request of requests.filter((entry) => entry.body)) {
      expect(request.url).toBe(endpoint);
      expect(new URLSearchParams(request.body).get('PRO_ITEM')).toBe('M00');
      expect(new URLSearchParams(request.body).get('SDATE')).toBe('115/09/26');
      expect(new URLSearchParams(request.body).get('EDATE')).toBe('115/09/26');
    }
    expect(requests).toContainEqual({ url: detail });
    expect(result).toMatchObject({ status: 'degraded', dataDate: '2026-09-26', warning: expect.stringContaining('可能漏筆') });
    expect(result.events).toHaveLength(2);
    expect(result.events[0]).toMatchObject({ market: 'TWSE', stockCode: '2330', sourceKey: 'mops:TWSE:2330:1150926:070004:2', sourceUrl: detail });
    expect(result.events[0].content).toContain('董事會通過測試用營運計畫');
    expect(result.events[1]).toMatchObject({ market: 'TPEX', stockCode: '6488' });
  });

  it('does not claim complete or zero success for a valid empty query', async () => {
    const provider = createMopsSearchMaterialProvider({
      async post() { return '\ufeff\r\n{"status":"success","data":[]}'; },
      async get() { throw new Error('unexpected detail request'); },
    });
    await expect(provider.fetchForDate('2026-09-27')).resolves.toMatchObject({ status: 'degraded', events: [] });
  });

  it('treats the official no-announcement response as a degraded empty observation, not a source failure', async () => {
    const provider = createMopsSearchMaterialProvider({
      async post() { return '\ufeff{"status":"fail","message":["查無公告資料"],"data":[]}'; },
      async get() { throw new Error('no detail pages for an empty day'); },
    });
    await expect(provider.fetchForDate('2026-09-28')).resolves.toMatchObject({
      status: 'degraded', dataDate: '2026-09-28', events: [], warning: expect.stringContaining('查無公告資料'),
    });
  });

  it('still rejects other failed website responses', async () => {
    const provider = createMopsSearchMaterialProvider({
      async post() { return { status: 'fail', message: '安全性封鎖', data: [] }; },
      async get() { throw new Error('unexpected detail request'); },
    });
    await expect(provider.fetchForDate('2026-09-28')).rejects.toThrow('未回傳成功資料列');
  });

  it('rejects another date or a non-official detail link', async () => {
    const provider = createMopsSearchMaterialProvider({
      async post() { return { status: 'success', data: [{ CDATE: '115/09/25', CTIME: '17:30:00', TYPEK: '上市', COMPANY_ID: '2330', COMPANY_NAME: '測試半導體', SUBJECT: '董事會決議', HYPERLINK: 'https://evil.example/steal' }] }; },
      async get() { throw new Error('unexpected detail request'); },
    });
    await expect(provider.fetchForDate('2026-09-26')).rejects.toThrow(/日期|連結/);
    const offOrigin = createMopsSearchMaterialProvider({
      async post() { return { status: 'success', data: [{
        CDATE: '115/09/26', CTIME: '07:00:04', TYPEK: '上市', COMPANY_ID: '2330',
        COMPANY_NAME: '測試半導體', SUBJECT: '董事會決議', HYPERLINK: 'https://evil.example/steal',
      }] }; },
      async get() { throw new Error('must not fetch an external detail link'); },
    });
    await expect(offOrigin.fetchForDate('2026-09-26')).rejects.toThrow(/不是官方來源/);
  });

  it('fails closed at the website result cap rather than claiming a complete date', async () => {
    const provider = createMopsSearchMaterialProvider({
      async post() { return { status: 'success', data: Array.from({ length: 1_000 }, () => ({})) }; },
      async get() { throw new Error('must not fetch capped results'); },
    });
    await expect(provider.fetchForDate('2026-09-26')).rejects.toThrow(/1000 筆上限/);
  });
});
