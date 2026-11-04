import { describe, expect, it } from 'vitest';
import { createMopsSearchMaterialProvider } from '../../src/providers/mops-search-material';
import { loadFixture } from '../support';

const endpoint = 'https://mopsov.twse.com.tw/mops/web/ezsearch_query';
const detail = 'https://mopsov.twse.com.tw/mops/web/ajax_t05sr01_1?firstin=true&step=1&SEQ_NO=2&SPOKE_TIME=70004&SPOKE_DATE=20260926&COMPANY_ID=2330';

function officialDetail(url: string): string {
  const query = new URL(url).searchParams;
  const code = query.get('COMPANY_ID');
  const sequence = query.get('SEQ_NO');
  return `<html><body><table id="table01"><tr><th>本資料由 (上市公司) ${code} 測試公司提供</th></tr><tr><th>序號</th><td>${sequence}</td></tr><tr><th>主旨</th><td>測試公告</td></tr><tr><th>說明</th><td>董事會通過測試用營運計畫</td></tr></table></body></html>`;
}

describe('material-event-monitoring / MOPS website search / official dated announcements', () => {
  it('queries both markets by ROC date, gets full detail, and warns that search results can miss daily records', async () => {
    const payload = JSON.parse(await loadFixture('providers/mops-search-material.json')) as { status: string; data: Array<{ TYPEK: string }> };
    const requests: Array<{ url: string; body?: string }> = [];
    const provider = createMopsSearchMaterialProvider({
      async post(url, body) {
        requests.push({ url, body });
        const label = new URLSearchParams(body).get('TYPEK') === 'sii' ? '上市' : '上櫃';
        return { ...payload, data: payload.data.filter((row) => row.TYPEK === label) };
      },
      async get(url) { requests.push({ url }); return officialDetail(url); },
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

  it('accepts an official foreign-issuer detail page with its own dated link format', async () => {
    const foreignDetail = 'https://mopsov.twse.com.tw/mops/web/ajax_t59sb01?firstin=true&TYPEK=all&DATE1=20261102&SKEY=2&step=2b&co_id=912000';
    const detailHtml = '<html><body><table><tr><th>公司代號</th><td>912000</td></tr><tr><th>公告序號</th><td>2</td></tr><tr><th>主旨</th><td>測試公告</td></tr><tr><th>發生事由</th><td>測試用完整外國發行人公告內容</td></tr></table></body></html>';
    const fetched: string[] = [];
    const provider = createMopsSearchMaterialProvider({
      async post(_url, body) {
        return { status: 'success', data: new URLSearchParams(body).get('TYPEK') === 'sii' ? [{
          CDATE: '115/11/02', CTIME: '16:54:14', TYPEK: '上市', COMPANY_ID: '912000',
          COMPANY_NAME: '測試外國發行人', SUBJECT: '測試公告', HYPERLINK: foreignDetail,
        }] : [] };
      },
      async get(url) { fetched.push(url); return detailHtml; },
    });

    const result = await provider.fetchForDate('2026-11-02');

    expect(fetched).toEqual([foreignDetail]);
    expect(result.events).toHaveLength(1);
    expect(result.events[0]).toMatchObject({
      stockCode: '912000', sourceUrl: foreignDetail,
      sourceKey: 'mops:TWSE:912000:1151102:165414:2',
      content: expect.stringContaining('測試用完整外國發行人公告內容'),
    });
  });

  it('keeps a valid announcement when another row has an unrecognized detail link', async () => {
    const payload = JSON.parse(await loadFixture('providers/mops-search-material.json')) as { data: Array<{ TYPEK: string; HYPERLINK: string }> };
    const requested: string[] = [];
    const provider = createMopsSearchMaterialProvider({
      async post(_url, body) {
        const label = new URLSearchParams(body).get('TYPEK') === 'sii' ? '上市' : '上櫃';
        return { status: 'success', data: payload.data.filter((row) => row.TYPEK === label).map((row) => ({
          ...row, HYPERLINK: label === '上櫃' ? 'https://evil.example/steal' : row.HYPERLINK,
        })) };
      },
      async get(url) {
        requested.push(url);
        return '<html><body><table><tr><th>本資料由 (上市公司) 2330 測試公司提供</th></tr><tr><th>序號</th><td>2</td></tr><tr><th>說明</th><td>測試用完整公告內容</td></tr></table></body></html>';
      },
    });

    const result = await provider.fetchForDate('2026-09-26');

    expect(requested).toEqual([detail]);
    expect(result.events.map((event) => event.stockCode)).toEqual(['2330']);
    expect(result.warning).toContain('6488');
    expect(result.warning).toContain('資料列驗證失敗');
  });

  it('rejects a same-host detail link on a nonstandard HTTPS port', async () => {
    const provider = createMopsSearchMaterialProvider({
      async post(_url, body) { return { status: 'success', data: new URLSearchParams(body).get('TYPEK') === 'sii' ? [{
        CDATE: '115/09/26', CTIME: '07:00:04', TYPEK: '上市', COMPANY_ID: '2330',
        COMPANY_NAME: '測試半導體', SUBJECT: '董事會決議',
        HYPERLINK: detail.replace('mopsov.twse.com.tw', 'mopsov.twse.com.tw:4444'),
      }] : [] }; },
      async get() { return '<html><body><table><tr><th>說明</th><td>不應擷取</td></tr></table></body></html>'; },
    });

    await expect(provider.fetchForDate('2026-09-26')).rejects.toThrow(/不是官方來源|資料列驗證失敗/);
  });

  it('rejects an official detail page whose table belongs to another company or sequence', async () => {
    const provider = createMopsSearchMaterialProvider({
      async post(_url, body) { return { status: 'success', data: new URLSearchParams(body).get('TYPEK') === 'sii' ? [{
        CDATE: '115/09/26', CTIME: '07:00:04', TYPEK: '上市', COMPANY_ID: '2330',
        COMPANY_NAME: '測試半導體', SUBJECT: '董事會決議', HYPERLINK: detail,
      }] : [] }; },
      async get() { return '<html><body><table><tr><th>本資料由 (上市公司) 9999 其他公司提供</th></tr><tr><th>序號</th><td>3</td></tr><tr><th>說明</th><td>其他公司的完整公告</td></tr></table></body></html>'; },
    });

    await expect(provider.fetchForDate('2026-09-26')).rejects.toThrow(/公告明細取得失敗/);
  });

  it('keeps other complete announcements when one official detail fails and flags the missing detail', async () => {
    const payload = JSON.parse(await loadFixture('providers/mops-search-material.json')) as { data: Array<{ TYPEK: string }> };
    const provider = createMopsSearchMaterialProvider({
      async post(_url, body) {
        const label = new URLSearchParams(body).get('TYPEK') === 'sii' ? '上市' : '上櫃';
        return { status: 'success', data: payload.data.filter((row) => row.TYPEK === label) };
      },
      async get(url) {
        if (url.includes('COMPANY_ID=2330')) throw new Error('Request failed with status code 404');
        return officialDetail(url);
      },
    });

    const result = await provider.fetchForDate('2026-09-26');

    expect(result).toMatchObject({ status: 'degraded', warning: expect.stringContaining('2330') });
    expect(result.events).toHaveLength(1);
    expect(result.events[0]).toMatchObject({ stockCode: '6488', content: expect.stringContaining('董事會通過測試用營運計畫') });
    expect(result.warning).toContain('明細取得失敗');
    expect(result.warning).toContain('HTTP 404');
  });

  it('retries a temporarily unavailable official detail after the first pass before omitting it', async () => {
    const payload = JSON.parse(await loadFixture('providers/mops-search-material.json')) as { data: Array<{ TYPEK: string }> };
    let attempts = 0;
    const provider = createMopsSearchMaterialProvider({
      async post(_url, body) {
        const label = new URLSearchParams(body).get('TYPEK') === 'sii' ? '上市' : '上櫃';
        return { status: 'success', data: payload.data.filter((row) => row.TYPEK === label) };
      },
      async get(url) {
        if (url.includes('COMPANY_ID=2330') && ++attempts === 1) throw new Error('Request failed with status code 404');
        return officialDetail(url);
      },
    });

    const result = await provider.fetchForDate('2026-09-26');

    expect(attempts).toBe(2);
    expect(result.events).toHaveLength(2);
    expect(result.events[0].content).toContain('董事會通過測試用營運計畫');
    expect(result.warning).not.toContain('明細取得失敗');
  });

  it('does not repeatedly request a detail page that returns an invalid HTML shape', async () => {
    const payload = JSON.parse(await loadFixture('providers/mops-search-material.json')) as { data: Array<{ TYPEK: string }> };
    let invalidAttempts = 0;
    const provider = createMopsSearchMaterialProvider({
      async post(_url, body) {
        const label = new URLSearchParams(body).get('TYPEK') === 'sii' ? '上市' : '上櫃';
        return { status: 'success', data: payload.data.filter((row) => row.TYPEK === label) };
      },
      async get(url) {
        if (url.includes('COMPANY_ID=2330')) { invalidAttempts += 1; return '<html><body>暫無有效明細</body></html>'; }
        return officialDetail(url);
      },
    });

    const result = await provider.fetchForDate('2026-09-26');

    expect(invalidAttempts).toBe(1);
    expect(result.events).toHaveLength(1);
    expect(result.warning).toContain('回應格式無效 1 筆');
  });

  it('fails instead of calling a day empty when every announcement detail is unavailable', async () => {
    const payload = JSON.parse(await loadFixture('providers/mops-search-material.json')) as { data: Array<{ TYPEK: string }> };
    const provider = createMopsSearchMaterialProvider({
      async post(_url, body) {
        const label = new URLSearchParams(body).get('TYPEK') === 'sii' ? '上市' : '上櫃';
        return { status: 'success', data: payload.data.filter((row) => row.TYPEK === label) };
      },
      async get() { throw new Error('Request failed with status code 404'); },
    });

    await expect(provider.fetchForDate('2026-09-26')).rejects.toThrow(/公告明細取得失敗.*2330.*6488.*HTTP 404/);
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
