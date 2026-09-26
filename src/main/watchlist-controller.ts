import { createRepositories } from '../repositories';
import type { CompanyDirectory } from '../services/watchlist';
import { addListedCompany, removeWatchlistCompany, setWatchlistCompanyActive, updateWatchlistCompany } from '../services/watchlist';
import { applyLegacyWatchlistImport, previewLegacyWatchlistImport } from '../services/watchlist-import';

type Repositories = ReturnType<typeof createRepositories>;

export function createWatchlistController(dependencies: {
  repositories: Repositories;
  directory: CompanyDirectory;
  now?: () => string;
}) {
  const now = dependencies.now ?? (() => new Date().toISOString());
  const { repositories, directory } = dependencies;

  return {
    list(filter: { query?: string; activeOnly?: boolean } = {}) {
      return repositories.watchlist.list(filter);
    },
    add(input: { market: 'TWSE' | 'TPEX'; stockCode: string; category?: string; notes?: string }) {
      return addListedCompany({
        directory,
        now,
        repository: {
          transaction: repositories.transaction,
          saveCompany(company) {
            const { id: _unusedDomainId, ...data } = company;
            return repositories.companies.upsert(data).id;
          },
          saveWatchlist: repositories.watchlist.upsert,
        },
      })(input);
    },
    update(input: { companyId: string; category: string; notes: string }) {
      return updateWatchlistCompany(repositories.watchlist, input.companyId, input, now());
    },
    setActive(input: { companyId: string; active: boolean }) {
      return setWatchlistCompanyActive(repositories.watchlist, input.companyId, input.active, now());
    },
    remove(input: { companyId: string }) {
      return removeWatchlistCompany(repositories.watchlist, input.companyId, now());
    },
    previewImport(input: { configText: string }) {
      return previewLegacyWatchlistImport(input.configText, {
        directory,
        isWatched: repositories.watchlist.isWatched,
      });
    },
    applyImport(input: { configText: string }) {
      return applyLegacyWatchlistImport(input.configText, {
        directory,
        now,
        repository: {
          transaction: repositories.transaction,
          isWatched: repositories.watchlist.isWatched,
          saveCompany(listing, updatedAt) {
            return repositories.companies.upsert({ ...listing, updatedAt }).id;
          },
          saveWatchlist(entry) {
            return repositories.watchlist.upsert({
              companyId: entry.companyId,
              active: entry.active,
              category: entry.category,
              notes: entry.notes,
              createdAt: entry.createdAt,
              updatedAt: entry.updatedAt,
            });
          },
        },
      });
    },
  };
}
