import { createHash } from 'node:crypto';
import type { Market } from '../domain/watchlist';
import type { HttpClientPort, MaterialProviderResult, MaterialSourceRecord } from './mops-material';

type RecordValue = Record<string, unknown>;

function recordValue(value: unknown): RecordValue {
  return value && typeof value === 'object' && !Array.isArray(value) ? value as RecordValue : {};
}

function pick(record: RecordValue, aliases: string[]): string {
  const normalized = new Map(Object.entries(record).map(([key, value]) => [key.replace(/[\s_]/g, '').toLowerCase(), value]));
  for (const alias of aliases) {
    const value = normalized.get(alias.replace(/[\s_]/g, '').toLowerCase());
    if (typeof value === 'string' || typeof value === 'number') {
      const text = String(value).trim();
      if (text) return text;
    }
  }
  return '';
}

function normalizeDate(value: string): string {
  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  const roc = /^(\d{2,3})\/?(\d{2})\/?(\d{2})$/.exec(value.replace(/\./g, '/'))
    ?? /^(\d{2,3})\/(\d{1,2})\/(\d{1,2})$/.exec(value);
  if (iso) return value;
  if (roc) return `${String(Number(roc[1]) + 1911).padStart(4, '0')}-${roc[2].padStart(2, '0')}-${roc[3].padStart(2, '0')}`;
  throw new Error(`結構化重大訊息日期格式無效：${value}`);
}

function normalizeInstant(date: string, rawTime: string): string {
  const time = rawTime || '00:00:00';
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(time);
  if (!match) throw new Error(`結構化重大訊息時間格式無效：${time}`);
  const [year, month, day] = date.split('-').map(Number);
  const [hour, minute, second = '00'] = match.slice(1);
  const instant = new Date(Date.UTC(year, month - 1, day, Number(hour) - 8, Number(minute), Number(second)));
  if (Number(hour) > 23 || Number(minute) > 59 || Number(second) > 59) throw new Error(`結構化重大訊息時間超出有效範圍：${time}`);
  return instant.toISOString();
}

function payloadRows(payload: unknown): { rows: unknown[]; dataDate: string } {
  if (Array.isArray(payload)) return { rows: payload, dataDate: '' };
  const wrapper = recordValue(payload);
  const rows = Array.isArray(wrapper.data) ? wrapper.data : Array.isArray(wrapper.records) ? wrapper.records : null;
  if (!rows) throw new Error('結構化重大訊息回應不是資料列');
  const dataDate = pick(wrapper, ['dataDate', 'queryDate', '資料日期', '查詢日期']);
  return { rows, dataDate: dataDate ? normalizeDate(dataDate) : '' };
}

const rowDateAliases = ['發言日期', '公告日期', '日期', 'Date', 'date'];
const codeAliases = ['公司代號', '股票代號', '股票代碼', '證券代號', 'stockCode', '公司代碼'];
const titleAliases = ['主旨', '公告主旨', 'subject', 'title'];

export function createMaterialReconciliationProvider(
  http: HttpClientPort,
  market: Market,
  endpoint = market === 'TWSE'
    ? 'https://openapi.twse.com.tw/v1/opendata/t187ap04_L'
    : 'https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap04_O',
) {
  return {
    async fetchForDate(targetDate: string): Promise<MaterialProviderResult> {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(targetDate)) throw new Error('查詢日期必須使用 YYYY-MM-DD 格式');
      const { rows, dataDate: responseDate } = payloadRows(await http.get(endpoint));
      const normalized = rows.map((raw) => {
        const row = recordValue(raw);
        const date = normalizeDate(pick(row, rowDateAliases));
        const stockCode = pick(row, codeAliases);
        const companyName = pick(row, ['公司名稱', '公司簡稱', 'CompanyName', 'CompanyAbbreviation']) || stockCode;
        const title = pick(row, titleAliases);
        const time = pick(row, ['發言時間', '公告時間', '時間', 'Time', 'time']);
        const content = pick(row, ['說明', '內容', '公告內容', 'content', 'description']) || title;
        if (!/^\d{4,6}$/.test(stockCode) || !title || !content) throw new Error(`${market} 結構化重大訊息缺少股票代號、主旨或內容`);
        const sequence = pick(row, ['序號', '流水號', '編號', 'seq', 'id']);
        const sourceUrl = pick(row, ['連結', '網址', 'link', 'url']) || `https://mops.twse.com.tw/mops/web/t05sr01_1?stock=${stockCode}`;
        const fingerprint = createHash('sha256').update(`${market}|${stockCode}|${date}|${time}|${title}|${content}`).digest('hex');
        const sourceKey = sequence ? `${market.toLowerCase()}:${sequence}` : `${market.toLowerCase()}:${fingerprint}`;
        const event: MaterialSourceRecord = {
          source: market === 'TWSE' ? 'twse' : 'tpex', market, stockCode, companyName,
          publishedAt: normalizeInstant(date, time), title, content, sourceKey, sourceUrl, revisionOf: null,
        };
        return { date, event };
      });
      const dataDate = responseDate || normalized.map(({ date }) => date).sort().at(-1) || '';
      if (!dataDate) throw new Error(`${market} 結構化重大訊息缺少權威資料日期`);
      return {
        status: 'complete',
        dataDate,
        events: normalized.filter(({ date }) => date === targetDate).map(({ event }) => event),
      };
    },
  };
}
