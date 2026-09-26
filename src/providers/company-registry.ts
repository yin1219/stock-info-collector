import type { Market } from '../domain/watchlist';

export interface CompanyListing {
  market: Market;
  stockCode: string;
  name: string;
}

export interface RegistryHttpClient {
  get(url: string): Promise<unknown>;
}

const codeKeys = ['公司代號', '公司代碼', '股票代號', '股票代碼', '證券代號', '證券代碼', 'SecuritiesCompanyCode', 'StockCode', 'CompanyCode'];
const nameKeys = ['公司名稱', '公司簡稱', '公司', 'CompanyName', 'CompanyAbbreviation', 'Name'];

function pickField(record: Record<string, unknown>, aliases: readonly string[]): string | undefined {
  const fields = new Map(Object.entries(record).map(([key, value]) => [key.replace(/[\s_]/g, '').toLowerCase(), value]));
  for (const alias of aliases) {
    const value = fields.get(alias.replace(/[\s_]/g, '').toLowerCase());
    if (typeof value === 'string' || typeof value === 'number') {
      const normalized = String(value).trim();
      if (normalized) return normalized;
    }
  }
  return undefined;
}

export function parseCompanyRegistry(payload: unknown, market: Market): CompanyListing[] {
  const rows = Array.isArray(payload)
    ? payload
    : payload && typeof payload === 'object' && 'data' in payload && Array.isArray(payload.data)
      ? payload.data
      : null;
  if (!rows) throw new Error(`${market} 公司名錄回應格式無效`);

  const listings = rows.flatMap((row): CompanyListing[] => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return [];
    const stockCode = pickField(row as Record<string, unknown>, codeKeys);
    const name = pickField(row as Record<string, unknown>, nameKeys);
    if (!stockCode || !/^\d{4,6}$/.test(stockCode) || !name) return [];
    return [{ market, stockCode, name }];
  });
  if (rows.length > 0 && listings.length === 0) {
    throw new Error(`${market} 公司名錄沒有可辨識的公司資料`);
  }
  return listings;
}

export function createCompanyRegistryProvider(http: RegistryHttpClient, endpoints: Record<Market, string> = {
  TWSE: 'https://openapi.twse.com.tw/v1/opendata/t187ap03_L',
  TPEX: 'https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O',
}) {
  const cache = new Map<Market, Promise<CompanyListing[]>>();

  async function getMarket(market: Market): Promise<CompanyListing[]> {
    let loading = cache.get(market);
    if (!loading) {
      loading = http.get(endpoints[market]).then((payload) => parseCompanyRegistry(payload, market));
      cache.set(market, loading);
      loading.catch(() => cache.delete(market));
    }
    return loading;
  }

  return {
    async findByStockCode(market: Market, stockCode: string): Promise<CompanyListing | undefined> {
      const normalizedCode = stockCode.trim();
      if (!/^\d{4,6}$/.test(normalizedCode)) return undefined;
      return (await getMarket(market)).find((listing) => listing.stockCode === normalizedCode);
    },
    async list(market: Market): Promise<CompanyListing[]> {
      return [...await getMarket(market)];
    },
    clearCache(market?: Market): void {
      if (market) cache.delete(market);
      else cache.clear();
    },
  };
}
