import { describe, expect, it } from 'vitest';
import { createMopsMaterialRssProvider } from '../../src/providers/mops-material-rss';

function big5Feed(itemCount = 2): Buffer {
  const items = Array.from({ length: itemCount }, (_, index) => Buffer.concat([
    Buffer.from(`<item><guid>rss-${index}</guid><stockCode>2330</stockCode><companyName>Test Company</companyName><market>TWSE</market><title>2330 `),
    Buffer.from([0xA4, 0xA4, 0xA4, 0xE5]),
    Buffer.from(` announcement</title><link>https://mops.example/${index}</link><pubDate>Sat, 26 Sep 2026 17:30:00 +0800</pubDate><description>Fixture description</description></item>`),
  ]));
  return Buffer.concat([Buffer.from('<rss><channel>'), ...items, Buffer.from('</channel></rss>')]);
}

describe('material-event-monitoring / provider contract / Big5 RSS fallback', () => {
  it('decodes Big5 and limits its degraded result to the configured item cap', async () => {
    const requests: string[] = [];
    const provider = createMopsMaterialRssProvider({ async get(url) { requests.push(url); return big5Feed(3); } }, 'fixture://mops/rss', 2);
    const result = await provider.fetchForDate('2026-09-26');
    expect(requests).toEqual(['fixture://mops/rss']);
    expect(result.status).toBe('degraded');
    expect(result.events).toHaveLength(2);
    expect(result.events[0]).toMatchObject({ stockCode: '2330', title: '2330 中文 announcement', content: 'Fixture description' });
  });

  it('rejects malformed feeds instead of treating them as an empty result', async () => {
    const provider = createMopsMaterialRssProvider({ async get() { return '<html>維護中</html>'; } }, 'fixture://mops/rss');
    await expect(provider.fetchForDate('2026-09-26')).rejects.toThrow(/RSS/);
  });
});
