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
  const compact = /^(\d{4})(\d{2})(\d{2})$/.exec(value);
  if (compact) return `${compact[1]}-${compact[2]}-${compact[3]}`;
  const roc = /^(\d{2,3})[/.](\d{1,2})[/.](\d{1,2})$/.exec(value);
  if (!roc) throw new Error(`違約交割資料日期格式無效：${value}`);
  return `${String(Number(roc[1]) + 1911).padStart(4, '0')}-${roc[2].padStart(2, '0')}-${roc[3].padStart(2, '0')}`;
}

function payload(payloadValue: unknown, market: Market): { dataDate: string; rows: unknown[] } {
  if (Array.isArray(payloadValue)) return { dataDate: '', rows: payloadValue };
  if (!payloadValue || typeof payloadValue !== 'object') throw new Error('違約交割來源回應格式無效');
  const wrapper = payloadValue as Record<string, unknown>;
  if (market === 'TWSE' && Array.isArray(wrapper.tables)) {
    if (wrapper.stat !== 'OK') throw new Error('TWSE 違約交割來源未回報成功');
    const summary = wrapper.tables.find((value) => value && typeof value === 'object'
      && Array.isArray((value as Record<string, unknown>).fields)
      && (value as { fields: unknown[] }).fields.includes('申報日期')
      && !(value as { fields: unknown[] }).fields.includes('證券代號')) as Record<string, unknown> | undefined;
    const security = wrapper.tables.find((value) => value && typeof value === 'object'
      && Array.isArray((value as Record<string, unknown>).fields)
      && (value as { fields: unknown[] }).fields.includes('證券代號')) as Record<string, unknown> | undefined;
    if (!summary || !Array.isArray(summary.data) || !summary.data.length
      || !security || !Array.isArray(security.fields) || !Array.isArray(security.data)) {
      throw new Error('TWSE 違約交割摘要申報日期或個股表格不存在');
    }
    const reportDates = summary.data.map((row) => {
      if (!Array.isArray(row) || typeof row[0] !== 'string') throw new Error('TWSE 違約交割摘要申報日期格式無效');
      return normalizeDate(row[0]);
    });
    const fields = security.fields as string[];
    const rows = security.data.flatMap((row) => {
      if (!Array.isArray(row) || row.length !== fields.length) throw new Error('TWSE 違約交割個股資料列格式無效');
      if (row[0] === '總計') return [];
      return [Object.fromEntries(fields.map((field, index) => [field, row[index]]))];
    });
    return { dataDate: reportDates.sort().at(-1)!, rows };
  }
  if (market === 'TPEX' && Array.isArray(wrapper.tables)) {
    if (wrapper.stat !== 'ok') throw new Error('TPEX 違約交割來源未回報成功');
    const table = wrapper.tables.find((value) => value && typeof value === 'object'
      && Array.isArray((value as Record<string, unknown>).fields)
      && ((value as { fields: unknown[] }).fields).includes('證券代號')) as Record<string, unknown> | undefined;
    if (!table || !Array.isArray(table.fields) || !Array.isArray(table.data)) {
      throw new Error('TPEX 違約交割個股表格不存在或格式無效');
    }
    const fields = table.fields as string[];
    if (typeof table.totalCount !== 'number' || table.totalCount !== table.data.length) {
      throw new Error('TPEX 違約交割個股表格筆數不符');
    }
    const dateRange = pick(wrapper, ['date']);
    const dataDate = normalizeDate(dateRange.split('~').at(-1) ?? '');
    const rows = table.data.map((value) => {
      if (!Array.isArray(value) || value.length !== fields.length) {
        throw new Error('TPEX 違約交割個股資料列格式無效');
      }
      return Object.fromEntries(fields.map((field, index) => [field, value[index]]));
    });
    return { dataDate, rows };
  }
  if (market === 'TWSE' && typeof wrapper.stat === 'string'
    && /沒有符合條件的資料/.test(wrapper.stat)
    && !pick(wrapper, ['dataDate', 'queryDate', 'date', '資料日期', '查詢日期', '日期'])) {
    throw new Error('TWSE 違約交割來源回報無符合資料，但未提供資料日期，不能視為本日零筆');
  }
  const date = pick(wrapper, ['dataDate', 'queryDate', 'date', '資料日期', '查詢日期', '日期']);
  const rows = Array.isArray(wrapper.stockDetail) ? wrapper.stockDetail
    : Array.isArray(wrapper.data) ? wrapper.data
      : Array.isArray(wrapper.records) ? wrapper.records
        : null;
  if (!rows) throw new Error('違約交割來源回應格式無效：缺少資料列');
  return { dataDate: date ? normalizeDate(date) : '', rows };
}

const dateAliases = ['date', 'disclosureDate', '資料日期', '揭露日期', '日期', '申報日期'];
const codeAliases = ['stockCode', '公司代號', '股票代號', '證券代號', '股票代碼'];
const brokerAliases = ['brokerCode', '券商代號', '證券商代號', '券商代碼', '證券商名稱'];
const twseDashboardEndpoint = 'https://www.twse.com.tw/rwd/zh/announcement/BFIGTU';

export function createDefaultDisclosureProvider(
  http: DisclosureHttpClient,
  market: Market,
  endpoint = market === 'TWSE'
    ? twseDashboardEndpoint
    : 'https://www.tpex.org.tw/www/zh-tw/bulletin/breach',
) {
  return {
    async fetchForDate(targetDate: string): Promise<DefaultDisclosureResult> {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(targetDate)) throw new Error('違約交割查詢日期必須使用 YYYY-MM-DD 格式');
      const requestUrl = endpoint === twseDashboardEndpoint ? (() => {
        const [year, month] = targetDate.split('-').map(Number);
        const startDate = new Date(Date.UTC(year, month - 2, 1)).toISOString().slice(0, 10).replace(/-/g, '');
        return `${endpoint}?startDate=${startDate}&endDate=${targetDate.replace(/-/g, '')}&response=json`;
      })() : endpoint;
      const { dataDate: responseDate, rows } = payload(await http.get(requestUrl), market);
      const normalized = rows.map((value) => {
        if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`${market} 違約交割資料列格式無效`);
        const row = value as Record<string, unknown>;
        const date = normalizeDate(pick(row, dateAliases) || responseDate);
        const stockCode = pick(row, codeAliases);
        const companyName = pick(row, ['companyName', '公司名稱', '公司簡稱', '公司', '證券名稱']) || stockCode;
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
