import { createCompany, type Company, type Market } from '../domain/watchlist';

export interface CompanyListing {
  market: Market;
  stockCode: string;
  name: string;
}

export interface CompanyDirectory {
  findByStockCode(market: Market, stockCode: string): Promise<CompanyListing | undefined>;
}

export interface WatchlistRepositoryPort {
  transaction<T>(work: () => T): T;
  saveCompany(company: Omit<Company, 'watchlist'> & { updatedAt: string }): string;
  saveWatchlist(entry: {
    companyId: string;
    active: boolean;
    category: string;
    notes: string;
    createdAt: string;
    updatedAt: string;
  }): unknown;
}

export interface WatchlistManagementRepository {
  updateDetails(companyId: string, details: { category: string; notes: string }, updatedAt: string): unknown;
  setActive(companyId: string, active: boolean, updatedAt: string): unknown;
  remove(companyId: string, updatedAt: string): unknown;
}

export function updateWatchlistCompany(
  repository: Pick<WatchlistManagementRepository, 'updateDetails'>,
  companyId: string,
  details: { category: string; notes: string },
  updatedAt: string,
): unknown {
  return repository.updateDetails(companyId, {
    category: details.category.trim(),
    notes: details.notes.trim(),
  }, updatedAt);
}

export function setWatchlistCompanyActive(
  repository: Pick<WatchlistManagementRepository, 'setActive'>,
  companyId: string,
  active: boolean,
  updatedAt: string,
): unknown {
  return repository.setActive(companyId, active, updatedAt);
}

export function removeWatchlistCompany(
  repository: Pick<WatchlistManagementRepository, 'remove'>,
  companyId: string,
  updatedAt: string,
): unknown {
  return repository.remove(companyId, updatedAt);
}

export function addListedCompany(dependencies: {
  directory: CompanyDirectory;
  repository: WatchlistRepositoryPort;
  now?: () => string;
}) {
  const now = dependencies.now ?? (() => new Date().toISOString());

  return async (input: { market: Market; stockCode: string; category?: string; notes?: string }) => {
    const stockCode = input.stockCode.trim();
    if (!/^\d{4,6}$/.test(stockCode)) {
      throw new Error('股票代號必須是 4 至 6 位數字');
    }

    const listing = await dependencies.directory.findByStockCode(input.market, stockCode);
    if (!listing) {
      throw new Error(`目前公司名錄找不到股票代號 ${stockCode}`);
    }
    if (listing.market !== input.market || listing.stockCode.trim() !== stockCode) {
      throw new Error('公司名錄回傳資料與查詢的市場或股票代號不符');
    }

    const company = createCompany(listing);
    const updatedAt = now();
    return dependencies.repository.transaction(() => {
      const companyId = dependencies.repository.saveCompany({
        id: company.id,
        market: company.market,
        stockCode: company.stockCode,
        name: company.name,
        updatedAt,
      });
      dependencies.repository.saveWatchlist({
        companyId,
        active: true,
        category: input.category?.trim() ?? '',
        notes: input.notes?.trim() ?? '',
        createdAt: updatedAt,
        updatedAt,
      });
      return { companyId, market: company.market, stockCode: company.stockCode, name: company.name, active: true };
    });
  };
}
