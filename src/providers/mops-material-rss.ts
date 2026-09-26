import { createHash } from 'node:crypto';
import * as cheerio from 'cheerio';
import type { MaterialProviderResult, MaterialSourceRecord, HttpClientPort } from './mops-material';

function feedText(value: unknown): string {
  if (typeof value === 'string') return value;
  if (Buffer.isBuffer(value)) return new TextDecoder('big5', { fatal: true }).decode(value);
  if (value && typeof value === 'object' && 'data' in value) return feedText(value.data);
  throw new Error('MOPS RSS 回應不是 Big5 文字或位元組');
}

function childText($: cheerio.CheerioAPI, item: cheerio.AnyNode, aliases: string[]): string {
  for (const alias of aliases) {
    const value = $(item).children(alias).first().text().trim();
    if (value) return value;
  }
  return '';
}

export function createMopsMaterialRssProvider(
  http: HttpClientPort,
  endpoint = 'https://mops.twse.com.tw/nas/rss/mopsrss201001.xml',
  itemLimit = 50,
) {
  if (!Number.isSafeInteger(itemLimit) || itemLimit < 1) throw new RangeError('RSS 項目上限必須是正整數');
  return {
    async fetchForDate(targetDate: string): Promise<MaterialProviderResult> {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(targetDate)) throw new Error('查詢日期必須使用 YYYY-MM-DD 格式');
      const xml = feedText(await http.get(endpoint));
      const $ = cheerio.load(xml, { xmlMode: true });
      if (!$('rss > channel').length && !$('feed').length) throw new Error('MOPS RSS 結構無效');
      const items = $('channel > item, feed > entry').toArray().slice(0, itemLimit);
      const events: MaterialSourceRecord[] = items.map((item) => {
        const title = childText($, item, ['title']);
        const stockCode = childText($, item, ['stockCode', 'StockCode', '公司代號', '股票代號'])
          || /\b(\d{4,6})\b/.exec(title)?.[1] || '';
        const companyName = childText($, item, ['companyName', 'CompanyName', '公司名稱']) || stockCode;
        const marketValue = childText($, item, ['market', 'Market', '市場別']);
        const market = marketValue === 'TPEX' || marketValue.includes('上櫃') ? 'TPEX' : 'TWSE';
        const sourceUrl = childText($, item, ['link']) || $(item).find('link').first().attr('href') || '';
        const publishedText = childText($, item, ['pubDate', 'published', 'updated']);
        const publishedAt = new Date(publishedText).toISOString();
        if (!/^\d{4,6}$/.test(stockCode) || !title || !sourceUrl) throw new Error('MOPS RSS 項目缺少股票代號、主旨或來源連結');
        const content = childText($, item, ['description', 'summary', 'content:encoded']) || title;
        const sourceId = childText($, item, ['guid', 'id']) || sourceUrl;
        const sourceKey = `mops-rss:${createHash('sha256').update(sourceId).digest('hex')}`;
        return {
          source: 'mops', market, stockCode, companyName, publishedAt, title, content,
          sourceKey, sourceUrl, revisionOf: null,
        };
      });
      return { status: 'degraded', dataDate: targetDate, events };
    },
  };
}
