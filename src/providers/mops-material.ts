import { createHash } from 'node:crypto';
import * as cheerio from 'cheerio';

export interface MaterialSourceRecord {
  source: 'mops' | 'twse' | 'tpex';
  market: 'TWSE' | 'TPEX';
  stockCode: string;
  companyName: string;
  publishedAt: string;
  title: string;
  content: string;
  sourceKey: string;
  sourceUrl: string;
  revisionOf: string | null;
}

export interface MaterialProviderResult {
  status: 'complete' | 'degraded';
  dataDate: string;
  events: MaterialSourceRecord[];
}

export interface HttpClientPort {
  get(url: string): Promise<unknown>;
}

function parseTaipeiInstant(rocDate: string, time: string): { iso: string; compactDate: string; compactTime: string } {
  const dateMatch = /^(\d{2,3})\/(\d{1,2})\/(\d{1,2})$/.exec(rocDate.trim());
  const timeMatch = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(time.trim());
  if (!dateMatch || !timeMatch) throw new Error('重大訊息日期或時間格式無效');

  const [, rocYear, monthText, dayText] = dateMatch;
  const [, hourText, minuteText, secondText = '00'] = timeMatch;
  const year = Number(rocYear) + 1911;
  const month = Number(monthText);
  const day = Number(dayText);
  const hour = Number(hourText);
  const minute = Number(minuteText);
  const second = Number(secondText);
  const localDate = new Date(Date.UTC(year, month - 1, day));
  if (localDate.getUTCFullYear() !== year || localDate.getUTCMonth() !== month - 1 || localDate.getUTCDate() !== day
    || hour > 23 || minute > 59 || second > 59) {
    throw new Error('重大訊息日期或時間超出有效範圍');
  }

  const utc = new Date(Date.UTC(year, month - 1, day, hour - 8, minute, second));
  const compactDate = `${rocYear.padStart(3, '0')}${String(month).padStart(2, '0')}${String(day).padStart(2, '0')}`;
  const compactTime = `${String(hour).padStart(2, '0')}${String(minute).padStart(2, '0')}${String(second).padStart(2, '0')}`;
  return { iso: utc.toISOString(), compactDate, compactTime };
}

function headerIndex(headers: string[], candidates: readonly string[], label: string): number {
  const index = headers.findIndex((header) => candidates.includes(header.replace(/[\s　]/g, '')));
  if (index < 0) throw new Error(`重大訊息表格缺少必要欄位：${label}`);
  return index;
}

export function parseMopsMaterialHtml(html: string, baseUrl = 'https://mops.twse.com.tw'): MaterialSourceRecord[] {
  const $ = cheerio.load(html);
  const table = $('#table01').length ? $('#table01').first() : $('table').first();
  if (table.length === 0) throw new Error('重大訊息表格不存在，不能視為零筆成功');

  const headerRow = table.find('thead tr').first().length
    ? table.find('thead tr').first()
    : table.find('tr').first();
  const headers = headerRow.find('th,td').toArray().map((cell) => $(cell).text().trim().replace(/[\s　]/g, ''));
  if (headers.length === 0) throw new Error('重大訊息表格沒有欄位標題');

  const columns = {
    date: headerIndex(headers, ['發言日期', '公告日期', '日期'], '日期'),
    time: headerIndex(headers, ['發言時間', '公告時間', '時間'], '時間'),
    stockCode: headerIndex(headers, ['股票代號', '公司代號', '代號'], '股票代號'),
    companyName: headerIndex(headers, ['公司名稱', '公司'], '公司名稱'),
    market: headerIndex(headers, ['市場別', '市場'], '市場別'),
    title: headerIndex(headers, ['主旨', '公告主旨', '說明'], '主旨'),
  };

  const rows = table.find('tbody tr').length ? table.find('tbody tr') : table.find('tr').slice(1);
  return rows.toArray().map((row) => {
    const cells = $(row).find('td').toArray().map((cell) => $(cell).text().trim());
    const stockCode = cells[columns.stockCode] ?? '';
    const companyName = cells[columns.companyName] ?? '';
    const title = cells[columns.title] ?? '';
    const marketText = cells[columns.market] ?? '';
    if (!/^\d{4,6}$/.test(stockCode) || !companyName || !title) {
      throw new Error('重大訊息資料列缺少有效的股票代號、公司名稱或主旨');
    }

    const market = marketText === '上市' || marketText === 'TWSE'
      ? 'TWSE'
      : marketText === '上櫃' || marketText === 'TPEX'
        ? 'TPEX'
        : null;
    if (!market) throw new Error(`重大訊息資料列市場別無法辨識：${marketText || '(空白)'}`);

    const { iso, compactDate, compactTime } = parseTaipeiInstant(cells[columns.date] ?? '', cells[columns.time] ?? '');
    const link = $(row).find('a[href]').first();
    const href = link.attr('href') ?? '';
    const sourceUrl = href ? new URL(href, baseUrl).href : '';
    const query = new URL(sourceUrl).searchParams;
    const sequence = query.get('seq') ?? query.get('h1205');
    const fingerprint = createHash('sha256').update(`${market}|${stockCode}|${iso}|${title.replace(/\s+/g, ' ').trim()}`).digest('hex');
    const sourceKey = `mops:${market}:${stockCode}:${compactDate}:${compactTime}:${sequence ?? fingerprint}`;

    return {
      source: 'mops',
      market,
      stockCode,
      companyName,
      publishedAt: iso,
      title,
      content: title,
      sourceKey,
      sourceUrl,
      revisionOf: query.get('revises'),
    };
  });
}

function responseHtml(response: unknown): string {
  return typeof response === 'string'
    ? response
    : response && typeof response === 'object' && 'data' in response && typeof response.data === 'string'
      ? response.data
      : '';
}

export function parseMopsMaterialDetail(html: string): string {
  const $ = cheerio.load(html);
  $('script, style, noscript').remove();
  const rows = $('#table01 tr, table tr').toArray().map((row) => {
    const cells = $(row).find('th,td').toArray().map((cell) => $(cell).text().replace(/\s+/g, ' ').trim());
    const label = cells[0] ?? '';
    if (/主旨|公告日期|發言日期|發言時間/.test(label)) return '';
    return cells.slice(1).filter(Boolean).join(' ');
  }).filter(Boolean);
  const content = rows.join('\n').trim() || $('body').text().replace(/\s+/g, ' ').trim();
  if (!content) throw new Error('重大訊息明細沒有公告內容');
  return content;
}

export function createMopsMaterialProvider(http: HttpClientPort, endpoint = 'https://mops.twse.com.tw/mops/web/ajax_t05st02') {
  return {
    async fetchForDate(targetDate: string): Promise<MaterialProviderResult> {
      const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(targetDate);
      if (!match) throw new Error('查詢日期必須使用 YYYY-MM-DD 格式');
      const rocYear = String(Number(match[1]) - 1911);
      const url = new URL(endpoint);
      url.searchParams.set('year', rocYear);
      url.searchParams.set('month', match[2]);
      url.searchParams.set('day', match[3]);
      const response = await http.get(url.href);
      const html = responseHtml(response);
      if (!html) throw new Error('MOPS 重大訊息回應不是可解析的 HTML');
      const endpointUrl = new URL(endpoint);
      const events = parseMopsMaterialHtml(html, `${endpointUrl.protocol}//${endpointUrl.host}/`);
      for (const event of events) {
        if (!event.sourceUrl) throw new Error(`重大訊息缺少明細連結：${event.stockCode}`);
        const detailHtml = responseHtml(await http.get(event.sourceUrl));
        if (!detailHtml) throw new Error(`重大訊息明細回應不是 HTML：${event.stockCode}`);
        event.content = parseMopsMaterialDetail(detailHtml);
      }
      return { status: 'complete', dataDate: targetDate, events };
    },
  };
}
