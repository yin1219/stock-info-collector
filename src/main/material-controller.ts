import { createRepositories } from '../repositories';

export function createMaterialController(repositories: ReturnType<typeof createRepositories>, now = () => new Date().toISOString()) {
  return {
    list(filter: { query?: string; unreadOnly?: boolean; watchedOnly?: boolean; eventIds?: string[] }) {
      return repositories.materialEvents.list(filter);
    },
    detail(id: string) {
      const detail = repositories.materialEvents.details(id);
      if (!detail) throw new Error('找不到指定的重大訊息');
      return detail;
    },
    markRead(id: string) {
      return repositories.materialEvents.markRead(id, now());
    },
    monitorStatus() {
      const latest = repositories.jobRuns.list('material-event-monitoring')[0];
      const reconciliation = repositories.jobRuns.list('material-event-reconciliation')[0];
      if (!latest && !reconciliation) return null;
      const reconciliationChecks = reconciliation ? repositories.sourceChecks.list(reconciliation.id) : [];
      const status = latest?.status === 'failed' || reconciliation?.status === 'failed' ? 'failed'
        : latest?.status === 'degraded' || reconciliation?.status === 'degraded' ? 'degraded'
          : latest?.status ?? reconciliation?.status ?? 'idle';
      return {
        status,
        startedAt: latest?.startedAt ?? reconciliation?.startedAt ?? null,
        finishedAt: latest?.finishedAt ?? reconciliation?.finishedAt ?? null,
        errorMessage: [latest?.errorMessage, reconciliation?.errorMessage].filter(Boolean).join('; ') || null,
        sources: [...(latest ? repositories.sourceChecks.list(latest.id) : []), ...reconciliationChecks],
      };
    },
  };
}
