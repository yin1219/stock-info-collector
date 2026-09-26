import type { CompanyListing } from '../providers/company-registry';
import type { CompanyDirectory } from './watchlist';

export interface WatchlistImportItem {
  stockCode: string;
  status: 'add' | 'skip' | 'failed';
  reason: string;
  listing?: CompanyListing;
}

export interface WatchlistImportReport {
  items: WatchlistImportItem[];
  summary: { added: number; skipped: number; failed: number };
}

export interface WatchlistImportRepository {
  transaction<T>(work: () => T): T;
  isWatched(market: string, stockCode: string): boolean;
  saveCompany(listing: CompanyListing, updatedAt: string): string;
  saveWatchlist(entry: {
    companyId: string;
    market: CompanyListing['market'];
    stockCode: string;
    active: true;
    category: string;
    notes: string;
    createdAt: string;
    updatedAt: string;
  }): unknown;
}

export function parseLegacyStockNumbers(configText: string): string[] {
  let parsed: unknown;
  try {
    parsed = JSON.parse(configText);
  } catch {
    throw new Error('既有設定不是有效的 JSON');
  }
  if (!parsed || typeof parsed !== 'object' || !('StockNumbers' in parsed) || typeof parsed.StockNumbers !== 'string') {
    throw new Error('既有設定缺少 StockNumbers 字串');
  }
  return parsed.StockNumbers.split(',').map((code) => code.trim()).filter(Boolean);
}

function summarize(items: WatchlistImportItem[]): WatchlistImportReport {
  return {
    items,
    summary: {
      added: items.filter(({ status }) => status === 'add').length,
      skipped: items.filter(({ status }) => status === 'skip').length,
      failed: items.filter(({ status }) => status === 'failed').length,
    },
  };
}

export async function previewLegacyWatchlistImport(configText: string, dependencies: {
  directory: CompanyDirectory;
  isWatched: (market: CompanyListing['market'], stockCode: string) => boolean | Promise<boolean>;
}): Promise<WatchlistImportReport> {
  const seen = new Set<string>();
  const items: WatchlistImportItem[] = [];
  for (const stockCode of parseLegacyStockNumbers(configText)) {
    if (!/^\d{4,6}$/.test(stockCode)) {
      items.push({ stockCode, status: 'failed', reason: '股票代號必須是 4 至 6 位數字' });
      continue;
    }
    if (seen.has(stockCode)) {
      items.push({ stockCode, status: 'skip', reason: '匯入清單中已有相同代號' });
      continue;
    }
    seen.add(stockCode);

    const lookups = await Promise.allSettled([
      dependencies.directory.findByStockCode('TWSE', stockCode),
      dependencies.directory.findByStockCode('TPEX', stockCode),
    ]);
    const listing = lookups.find((result) => result.status === 'fulfilled' && result.value)?.status === 'fulfilled'
      ? (lookups.find((result) => result.status === 'fulfilled' && result.value) as PromiseFulfilledResult<CompanyListing>).value
      : undefined;
    if (!listing) {
      const failures = lookups.filter((result): result is PromiseRejectedResult => result.status === 'rejected');
      items.push({
        stockCode,
        status: 'failed',
        reason: failures.length > 0 ? `公司名錄查詢失敗：${String(failures[0].reason)}` : '公司名錄找不到此股票代號',
      });
      continue;
    }

    if (await dependencies.isWatched(listing.market, listing.stockCode)) {
      items.push({ stockCode, status: 'skip', reason: '公司已在關注清單中', listing });
      continue;
    }
    items.push({ stockCode, status: 'add', reason: '可新增', listing });
  }
  return summarize(items);
}

export async function applyLegacyWatchlistImport(configText: string, dependencies: {
  directory: CompanyDirectory;
  repository: WatchlistImportRepository;
  now?: () => string;
}): Promise<WatchlistImportReport> {
  const report = await previewLegacyWatchlistImport(configText, {
    directory: dependencies.directory,
    isWatched: (market, stockCode) => dependencies.repository.isWatched(market, stockCode),
  });
  const now = dependencies.now ?? (() => new Date().toISOString());
  const items = await Promise.all(report.items.map(async (item) => {
    if (item.status !== 'add' || !item.listing) return item;
    const updatedAt = now();
    try {
      const companyId = dependencies.repository.transaction(() => {
        const id = dependencies.repository.saveCompany(item.listing!, updatedAt);
        dependencies.repository.saveWatchlist({
          companyId: id,
          market: item.listing!.market,
          stockCode: item.listing!.stockCode,
          active: true,
          category: '',
          notes: '',
          createdAt: updatedAt,
          updatedAt,
        });
        return id;
      });
      return { ...item, reason: `已新增公司 ${companyId}` };
    } catch (error) {
      return { ...item, status: 'failed' as const, reason: error instanceof Error ? error.message : String(error) };
    }
  }));
  return summarize(items);
}
