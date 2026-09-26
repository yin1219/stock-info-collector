import { createRepositories } from '../repositories';

export function createDisclosureController(repositories: ReturnType<typeof createRepositories>) {
  return {
    list(filter: { disclosureDate?: string; market?: 'TWSE' | 'TPEX' }) {
      return repositories.defaultDisclosures.list(filter);
    },
  };
}
