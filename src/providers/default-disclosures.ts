import { createHash } from 'node:crypto';
import type { Market } from '../domain/watchlist';

export interface DefaultDisclosureRecord {
  market: Market;
  disclosureDate: string;
  stockCode: string;
  companyName: string;
  brokerCode: string;
  disclosedAt: string;
  sourceKey: string;
  content: Record<string, unknown>;
}

export interface DefaultDisclosureResult {
  market: Market;
  dataDate: string;
  records: DefaultDisclosureRecord[];
}

export interface DisclosureHttpClient {
  get(url: string): Promise<unknown>;
}

function normalizeKey(value: string): string {
  return value.replace(/[\s_]/g, '').toLowerCase();
}

function pick(row: Record<string, unknown>, aliases: string[]): string {
  const fields = new Map(Object.entries(row).map(([key, value]) => [normalizeKey(key), value]));
  for (const alias of aliases) {
    const value = fields.get(normalizeKey(alias));
    if (typeof value === 'string' || typeof value === 'number') {
      const text = String(value).trim();
      if (text) return text;
    }
  }
  return '';
}

function normalizeDate(value: string): string {
  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (iso) return value;
  const roc = /^(\d{2,3})\/(\d{1,2})\/(\d{1,2})$/.exec(value);
  if (!roc) throw new Error(`違約交割資料日期格式無效：${value}`);
  return `${String(Number(roc[1]) + 1911).padStart(4, '0')}-${roc[2].padStart(2, '0')}-${roc[3].padStart(2, '0')}`;
}

function payload(payloadValue: unknown): { dataDate: string; rows: unknown[] } {
  if (Array.isArray(payloadValue)) return { dataDate: '', rows: payloadValue };
  if (!payloadValue || typeof payloadValue !== 'object') throw new Error('違約交割來源回應格式無效');
  const wrapper = payloadValue as Record<string, unknown>;
  const date = pick(wrapper, ['dataDate', 'queryDate', 'date', '資料日期', '查詢日期', '日期']);
  const rows = Array.isArray(wrapper.stockDetail) ? wrapper.stockDetail
    : Array.isArray(wrapper.data) ? wrapper.data
      : Array.isArray(wrapper.records) ? wrapper.records
        : null;
  if (!rows) throw new Error('違約交割來源回應格式無效：缺少資料列');
  return { dataDate: date ? normalizeDate(date) : '', rows };
}

const dateAliases = ['date', 'disclosureDate', '資料日期', '揭露日期', '日期'];
const codeAliases = ['stockCode', '公司代號', '股票代號', '證券代號', '股票代碼'];
const brokerAliases = ['brokerCode', '券商代號', '證券商代號', '券商代碼'];

export function createDefaultDisclosureProvider(
  http: DisclosureHttpClient,
  market: Market,
  endpoint = market === 'TWSE'
    ? 'https://www.twse.com.tw/rwd/zh/announcement/BFIGTU?response=json'
    : 'https://www.tpex.org.tw/openapi/v1/violation',
) {
  return {
    async fetchForDate(targetDate: string): Promise<DefaultDisclosureResult> {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(targetDate)) throw new Error('違約交割查詢日期必須使用 YYYY-MM-DD 格式');
      const { dataDate: responseDate, rows } = payload(await http.get(endpoint));
      const normalized = rows.map((value) => {
        if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`${market} 違約交割資料列格式無效`);
        const row = value as Record<string, unknown>;
        const date = normalizeDate(pick(row, dateAliases) || responseDate);
        const stockCode = pick(row, codeAliases);
        const companyName = pick(row, ['companyName', '公司名稱', '公司簡稱', '公司']) || stockCode;
        const brokerCode = pick(row, brokerAliases);
        if (!/^\d{4,6}$/.test(stockCode) || !companyName) throw new Error(`${market} 違約交割資料缺少有效股票代號或公司名稱`);
        const disclosedAt = `${date}T00:00:00+08:00`;
        const stableParts = [market, date, stockCode, brokerCode, pick(row, ['違約序號', '流水號', '序號', 'id'])];
        const sourceKey = `${market.toLowerCase()}:${createHash('sha256').update(stableParts.join('|')).digest('hex')}`;
        return {
          market, disclosureDate: date, stockCode, companyName, brokerCode, disclosedAt,
          sourceKey, content: row,
        } satisfies DefaultDisclosureRecord;
      });
      const dataDate = responseDate || normalized.map(({ disclosureDate }) => disclosureDate).sort().at(-1) || '';
      if (!dataDate) throw new Error(`${market} 違約交割回應缺少資料日期，不能判定為本日零筆`);
      return { market, dataDate, records: normalized.filter(({ disclosureDate }) => disclosureDate === targetDate) };
    },
  };
}
