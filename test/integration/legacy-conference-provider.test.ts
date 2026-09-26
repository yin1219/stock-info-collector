import { describe, expect, it } from 'vitest';
import { createLegacyConferenceProvider } from '../../src/providers/conference';
import { loadFixture } from '../support';

describe('conference-calendar-sync / provider contract / migrates the legacy MOPS fetch flow', () => {
  it('redirects each watchlist code and parses the resulting conference fixture', async () => {
    const html = await loadFixture('conferences/mops-conference.html');
    const calls: Array<{ method: string; url: string; body?: unknown }> = [];
    const http = {
      async post(url: string, body: unknown) {
        calls.push({ method: 'POST', url, body });
        return { result: { url: 'https://mops.twse.com.tw/mops/web/t100sb07_1?co_id=2454' } };
      },
      async get(url: string) { calls.push({ method: 'GET', url }); return html; },
    };
    const provider = createLegacyConferenceProvider(http, { redirectEndpoint: 'fixture://mops/redirect' });
    await expect(provider.fetch(['2454'])).resolves.toEqual([expect.objectContaining({
      CompId: '2454', CompName: '聯發科技股份有限公司', Location: '線上法人說明會',
    })]);
    expect(calls).toEqual([
      { method: 'POST', url: 'fixture://mops/redirect', body: { apiName: 'ajax_t100sb07_1', parameters: { co_id: '2454', encodeURIComponent: 1, step: 1, firstin: 1, off: 1, TYPEK: 'all' } } },
      { method: 'GET', url: 'https://mops.twse.com.tw/mops/web/t100sb07_1?co_id=2454' },
    ]);
  });

  it('rejects invalid codes and redirects outside the official MOPS host before fetching HTML', async () => {
    let fetchCount = 0;
    const provider = createLegacyConferenceProvider({
      async post() { return { result: { url: 'https://example.com/private' } }; },
      async get() { fetchCount += 1; return ''; },
    }, { redirectEndpoint: 'fixture://mops/redirect' });
    await expect(provider.fetch(['../token.json'])).rejects.toThrow(/股票代號/);
    await expect(provider.fetch(['2454'])).rejects.toThrow(/MOPS/);
    expect(fetchCount).toBe(0);
  });
});
