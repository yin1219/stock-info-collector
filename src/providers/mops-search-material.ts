import { createHash } from 'node:crypto';
import { parseMopsMaterialDetail, parseTaipeiInstant, type MaterialProviderResult, type MaterialSourceRecord } from './mops-material';

interface SearchHttpPort {
  post(url: string, body: string): Promise<unknown>;
  get(url: string): Promise<unknown>;
}

type SearchRow = Record<string, unknown>;

function searchRows(response: unknown): { rows: SearchRow[]; explicitlyEmpty: boolean } {
  const value = typeof response === 'string'
    ? JSON.parse(response.trim().replace(/^\uFEFF/, '')) as unknown
    : response;
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('MOPS 公告快易查回應格式無效');
  const result = value as { status?: unknown; message?: unknown; data?: unknown };
  const noAnnouncements = result.message === '查無公告資料'
    || (Array.isArray(result.message) && result.message.length === 1 && result.message[0] === '查無公告資料');
  if (result.status === 'fail' && noAnnouncements
    && Array.isArray(result.data) && result.data.length === 0) return { rows: [], explicitlyEmpty: true };
  if (result.status !== 'success' || !Array.isArray(result.data)) throw new Error('MOPS 公告快易查未回傳成功資料列');
  if (result.data.length >= 1_000) throw new Error('MOPS 公告快易查達 1000 筆上限，不能確認資料完整');
  return { rows: result.data as SearchRow[], explicitlyEmpty: false };
}

function text(row: SearchRow, key: string): string {
  return typeof row[key] === 'string' ? row[key].trim() : '';
}

function record(row: SearchRow, market: 'TWSE' | 'TPEX', targetDate: string): MaterialSourceRecord {
  const stockCode = text(row, 'COMPANY_ID');
  const companyName = text(row, 'COMPANY_NAME');
  const title = text(row, 'SUBJECT');
  const date = text(row, 'CDATE');
  const time = text(row, 'CTIME');
  const sourceUrl = text(row, 'HYPERLINK');
  if (!/^\d{4,6}$/.test(stockCode) || !companyName || !title || !sourceUrl) {
    throw new Error(`MOPS 公告快易查缺少有效代號、公司、主旨或明細連結：${stockCode || '(未知)'}`);
  }
  const parsed = parseTaipeiInstant(date, time);
  if (parsed.compactDate !== `${Number(targetDate.slice(0, 4)) - 1911}${targetDate.slice(5, 7)}${targetDate.slice(8, 10)}`) {
    throw new Error(`MOPS 公告快易查資料日期不符：${stockCode}`);
  }
  const url = new URL(sourceUrl);
  if (url.protocol !== 'https:' || url.hostname !== 'mopsov.twse.com.tw' || url.pathname !== '/mops/web/ajax_t05sr01_1') {
    throw new Error(`MOPS 公告快易查明細連結不是官方來源：${stockCode}`);
  }
  const sequence = url.searchParams.get('SEQ_NO')
    ?? createHash('sha256').update(`${stockCode}|${parsed.iso}|${title}`).digest('hex');
  if (url.searchParams.get('COMPANY_ID') !== stockCode
    || url.searchParams.get('SPOKE_DATE') !== targetDate.replace(/-/g, '')
    || url.searchParams.get('SPOKE_TIME')?.padStart(6, '0') !== parsed.compactTime) {
    throw new Error(`MOPS 公告快易查明細連結與公告欄位不符：${stockCode}`);
  }
  return {
    source: 'mops', market, stockCode, companyName, publishedAt: parsed.iso, title, content: title,
    sourceKey: `mops:${market}:${stockCode}:${parsed.compactDate}:${parsed.compactTime}:${sequence}`,
    sourceUrl: url.href, revisionOf: null,
  };
}

export function createMopsSearchMaterialProvider(
  http: SearchHttpPort,
  endpoint = 'https://mopsov.twse.com.tw/mops/web/ezsearch_query',
) {
  return {
    async fetchForDate(targetDate: string): Promise<MaterialProviderResult> {
      const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(targetDate);
      if (!match) throw new Error('重大訊息查詢日期必須使用 YYYY-MM-DD 格式');
      const rocDate = `${Number(match[1]) - 1911}/${match[2]}/${match[3]}`;
      const events: MaterialSourceRecord[] = [];
      const emptyMarkets: string[] = [];
      for (const [market, type, label] of [['TWSE', 'sii', '上市'], ['TPEX', 'otc', '上櫃']] as const) {
        const body = new URLSearchParams({
          step: '00', RADIO_CM: '1', TYPEK: type, CO_MARKET: '', CO_ID: '', PRO_ITEM: 'M00',
          SUBJECT: '', SDATE: rocDate, EDATE: rocDate, lang: 'TW', AN: '',
        }).toString();
        const { rows, explicitlyEmpty } = searchRows(await http.post(endpoint, body));
        if (explicitlyEmpty) emptyMarkets.push(label);
        if (rows.some((row) => text(row, 'TYPEK') !== label)) throw new Error(`MOPS 公告快易查市場別不符：${market}`);
        events.push(...rows.map((row) => record(row, market, targetDate)));
      }
      for (let offset = 0; offset < events.length; offset += 4) {
        await Promise.all(events.slice(offset, offset + 4).map(async (event) => {
          try {
            const response = await http.get(event.sourceUrl);
            const html = typeof response === 'string' ? response : response && typeof response === 'object' && 'data' in response ? response.data : '';
            if (typeof html !== 'string' || !html.includes('說明')) throw new Error('明細缺少說明');
            event.content = parseMopsMaterialDetail(html);
          } catch (error) {
            throw new Error(`MOPS 公告明細取得失敗：${event.stockCode}：${error instanceof Error ? error.message : String(error)}`);
          }
        }));
      }
      return {
        status: 'degraded', dataDate: targetDate, events,
        warning: `${emptyMarkets.length ? `${emptyMarkets.join('、')}官方回覆查無公告資料；` : ''}公告快易查可能漏筆，須以每日對帳補回，不能證明今日全無重大訊息`,
      };
    },
  };
}
