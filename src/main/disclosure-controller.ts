import { createRepositories } from '../repositories';

export function createDisclosureController(repositories: ReturnType<typeof createRepositories>) {
  return {
    list(filter: { disclosureDate?: string; market?: 'TWSE' | 'TPEX' }) {
      return repositories.defaultDisclosures.list(filter);
    },
    monitorStatus() {
      const latest = repositories.jobRuns.list('default-disclosure-monitoring')[0];
      if (!latest) return null;
      const summary = JSON.parse(latest.summaryJson) as { markets?: unknown };
      return {
        status: latest.status,
        startedAt: latest.startedAt,
        finishedAt: latest.finishedAt,
        errorMessage: latest.errorMessage,
        markets: summary.markets ?? {},
      };
    },
  };
}
