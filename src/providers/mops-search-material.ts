import { createHash } from 'node:crypto';
import * as cheerio from 'cheerio';
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

function detailFailureCategory(cause: unknown): string {
  const error = cause && typeof cause === 'object'
    ? cause as { code?: unknown; message?: unknown; response?: { status?: unknown } } : {};
  const message = typeof error.message === 'string' ? error.message : '';
  const status = typeof error.response?.status === 'number' ? error.response.status
    : Number(/Request failed with status code (\d{3})/.exec(message)?.[1]);
  if (status >= 400 && status <= 599) return `HTTP ${status}`;
  if (message.includes('maxContentLength')) return '內容大小限制';
  if (error.code === 'ECONNABORTED' || error.code === 'ETIMEDOUT') return '連線逾時';
  if (error.code === 'EAI_AGAIN') return 'DNS 暫時失敗';
  if (error.code === 'ECONNRESET') return '連線中斷';
  if (message.includes('明細缺少') || message.includes('重大訊息明細沒有公告內容')) return '回應格式無效';
  return '其他錯誤';
}

function rowFailureCategory(cause: unknown): string {
  const message = cause instanceof Error ? cause.message : '';
  if (message.includes('資料日期不符')) return '資料日期不符';
  if (message.includes('不是官方來源')) return '連結不是官方來源';
  if (message.includes('連結與公告欄位不符')) return '連結欄位不符';
  if (message.includes('市場別不符')) return '市場別不符';
  if (message.includes('缺少有效')) return '欄位缺漏';
  return '資料列無法解析';
}

function detailBelongsTo(event: MaterialSourceRecord, html: string): boolean {
  const $ = cheerio.load(html);
  const rows = $('table tr').toArray().map((row) => $(row).find('th,td').toArray()
    .map((cell) => $(cell).text().replace(/\s+/g, ' ').trim()));
  const field = (label: string) => rows.find((cells) => cells[0]?.replace(/[\s　]/g, '') === label)?.[1] ?? '';
  const url = new URL(event.sourceUrl);
  if (url.pathname === '/mops/web/ajax_t59sb01') {
    return field('公司代號') === event.stockCode
      && field('公告序號') === url.searchParams.get('SKEY')
      && rows.some((cells) => /^發生(依|事由)/.test(cells[0] ?? '') && cells.slice(1).some(Boolean));
  }
  const header = rows.find((cells) => cells[0]?.includes('本資料由'))?.[0] ?? '';
  return new RegExp(`(?:^|\\D)${event.stockCode}(?:\\D|$)`).test(header)
    && field('序號') === url.searchParams.get('SEQ_NO')
    && Boolean(field('說明'));
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
  const foreignIssuerDetail = url.pathname === '/mops/web/ajax_t59sb01';
  if (url.origin !== 'https://mopsov.twse.com.tw'
    || (!foreignIssuerDetail && url.pathname !== '/mops/web/ajax_t05sr01_1')) {
    throw new Error(`MOPS 公告快易查明細連結不是官方來源：${stockCode}`);
  }
  const sequence = url.searchParams.get(foreignIssuerDetail ? 'SKEY' : 'SEQ_NO')
    ?? createHash('sha256').update(`${stockCode}|${parsed.iso}|${title}`).digest('hex');
  const linkMatchesRow = foreignIssuerDetail
    ? url.searchParams.get('co_id') === stockCode
      && url.searchParams.get('DATE1') === targetDate.replace(/-/g, '')
      && url.searchParams.get('TYPEK') === 'all'
      && url.searchParams.get('step') === '2b'
      && /^\d+$/.test(url.searchParams.get('SKEY') ?? '')
    : url.searchParams.get('COMPANY_ID') === stockCode
      && url.searchParams.get('SPOKE_DATE') === targetDate.replace(/-/g, '')
      && url.searchParams.get('SPOKE_TIME')?.padStart(6, '0') === parsed.compactTime;
  if (!linkMatchesRow) {
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
      const invalidRows: Array<{ code: string; reason: string }> = [];
      for (const [market, type, label] of [['TWSE', 'sii', '上市'], ['TPEX', 'otc', '上櫃']] as const) {
        const body = new URLSearchParams({
          step: '00', RADIO_CM: '1', TYPEK: type, CO_MARKET: '', CO_ID: '', PRO_ITEM: 'M00',
          SUBJECT: '', SDATE: rocDate, EDATE: rocDate, lang: 'TW', AN: '',
        }).toString();
        const { rows, explicitlyEmpty } = searchRows(await http.post(endpoint, body));
        if (explicitlyEmpty) emptyMarkets.push(label);
        for (const row of rows) {
          try {
            if (text(row, 'TYPEK') !== label) throw new Error('市場別不符');
            events.push(record(row, market, targetDate));
          } catch (cause) {
            invalidRows.push({ code: text(row, 'COMPANY_ID') || '(未知)', reason: rowFailureCategory(cause) });
          }
        }
      }
      const invalidCodes = invalidRows.map(({ code }) => code);
      const invalidReasonCounts = new Map<string, number>();
      for (const { reason } of invalidRows) invalidReasonCounts.set(reason, (invalidReasonCounts.get(reason) ?? 0) + 1);
      const invalidReasonSummary = [...invalidReasonCounts].map(([reason, count]) => `${reason} ${count} 筆`).join('、');
      if (invalidRows.length && events.length === 0) {
        throw new Error(`MOPS 公告快易查資料列驗證失敗：${invalidCodes.slice(0, 10).join('、')}（共 ${invalidRows.length} 筆；${invalidReasonSummary}）；未保存當次網站公告`);
      }
      const failedDetails = new Set<MaterialSourceRecord>();
      const failureReasons = new Map<MaterialSourceRecord, string>();
      async function loadDetail(event: MaterialSourceRecord): Promise<void> {
        try {
          const response = await http.get(event.sourceUrl);
          const html = typeof response === 'string' ? response : response && typeof response === 'object' && 'data' in response ? response.data : '';
          if (typeof html !== 'string' || !detailBelongsTo(event, html)) throw new Error('明細缺少必要欄位或公告歸屬不符');
          event.content = parseMopsMaterialDetail(html);
          failedDetails.delete(event);
          failureReasons.delete(event);
        } catch (cause) {
          failedDetails.add(event);
          failureReasons.set(event, detailFailureCategory(cause));
        }
      }
      for (let offset = 0; offset < events.length; offset += 4) {
        await Promise.all(events.slice(offset, offset + 4).map(loadDetail));
      }
      const retryableReasons = new Set(['HTTP 404', '連線逾時', '連線中斷', 'DNS 暫時失敗']);
      for (const event of [...failedDetails].filter((item) => retryableReasons.has(failureReasons.get(item) ?? '')).slice(0, 8)) {
        await loadDetail(event);
      }
      const failedCodes = events.filter((event) => failedDetails.has(event)).map((event) => event.stockCode);
      const reasonCounts = new Map<string, number>();
      for (const event of failedDetails) {
        const reason = failureReasons.get(event) ?? '其他錯誤';
        reasonCounts.set(reason, (reasonCounts.get(reason) ?? 0) + 1);
      }
      const reasonSummary = [...reasonCounts].map(([reason, count]) => `${reason} ${count} 筆`).join('、');
      if (events.length && failedDetails.size === events.length) {
        throw new Error(`MOPS 公告明細取得失敗：${failedCodes.slice(0, 10).join('、')}（共 ${failedDetails.size} 筆；${reasonSummary}）；未保存當次網站公告`);
      }
      const detailWarning = failedDetails.size
        ? `${failedDetails.size} 筆公告明細取得失敗（${failedCodes.slice(0, 10).join('、')}；${reasonSummary}），未保存這些公告，稍後重試；`
        : '';
      const rowWarning = invalidRows.length
        ? `${invalidRows.length} 筆公告資料列驗證失敗（${invalidCodes.slice(0, 10).join('、')}；${invalidReasonSummary}），未保存這些公告；`
        : '';
      return {
        status: 'degraded', dataDate: targetDate, events: events.filter((event) => !failedDetails.has(event)),
        warning: `${rowWarning}${detailWarning}${emptyMarkets.length ? `${emptyMarkets.join('、')}官方回覆查無公告資料；` : ''}公告快易查可能漏筆，須以每日對帳補回，不能證明今日全無重大訊息`,
      };
    },
  };
}
