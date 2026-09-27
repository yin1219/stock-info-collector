import { prepareMaterialEvent } from '../domain/dedup';
import type { createRepositories } from '../repositories';
import type { Market } from '../domain/watchlist';
import type { MaterialProviderResult, MaterialSourceRecord } from '../providers/mops-material';
import { buildMaterialNotification } from './notification-messages';
import { deliverOutboxNotification, type NotificationChannel } from './notification-delivery';
import { shouldDeliverNotification } from './notification-outbox';

type Repositories = ReturnType<typeof createRepositories>;
type Provider = { fetchForDate(date: string): Promise<MaterialProviderResult> };
type SourceStatus = 'complete' | 'degraded' | 'stale' | 'failed';
type Result = { market: Market; status: SourceStatus; result?: MaterialProviderResult; error?: string };

export function createMaterialReconciliationMonitor(dependencies: {
  repositories: Repositories;
  providers: Record<Market, Provider>;
  channel: NotificationChannel;
  now?: () => string;
}) {
  const now = dependencies.now ?? (() => new Date().toISOString());
  return {
    async run(input: { targetDate: string; idempotencyKey: string }): Promise<{ jobRunId: string; status: 'complete' | 'degraded' | 'failed' | 'stale'; discovered: number }> {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(input.targetDate)) throw new Error('對帳查詢日期必須使用 YYYY-MM-DD 格式');
      const job = dependencies.repositories.transaction(() => dependencies.repositories.jobRuns.start({
        kind: 'material-event-reconciliation', idempotencyKey: input.idempotencyKey, startedAt: now(),
      }));
      if (!job.created) {
        const summary = JSON.parse(job.summaryJson || '{}') as { discovered?: number };
        return { jobRunId: job.id, status: job.status === 'active' ? 'degraded' : job.status, discovered: summary.discovered ?? 0 };
      }

      const results: Result[] = await Promise.all((['TWSE', 'TPEX'] as const).map(async (market) => {
        try {
          const result = await dependencies.providers[market].fetchForDate(input.targetDate);
          if (!/^\d{4}-\d{2}-\d{2}$/.test(result.dataDate) || result.dataDate > input.targetDate) {
            throw new Error(`${market} 對帳資料日期無效`);
          }
          return { market, result, status: result.dataDate < input.targetDate ? 'stale' as const : result.status };
        } catch (error) {
          const raw = error instanceof Error ? error.message : String(error);
          const safeError = raw.replace(/\bBearer\s+[^\s,;]+/gi, 'Bearer [REDACTED]');
          return { market, error: safeError, status: 'failed' as const };
        }
      }));
      const active = dependencies.repositories.watchlist.list({ activeOnly: true });
      const watched = new Set(active.map(({ market, stockCode }) => `${market}:${stockCode}`));
      const checkedAt = now();
      const discoveredIds: string[] = [];
      let notification: ReturnType<typeof buildMaterialNotification> = null;
      let dedupeKey: string | undefined;
      const status = results.every((item) => item.status === 'complete')
        ? 'complete'
        : results.some((item) => item.status === 'complete' || item.status === 'degraded')
          ? 'degraded' : results.some((item) => item.status === 'stale') ? 'stale' : 'failed';
      let deliverNotification = false;

      const persisted = dependencies.repositories.transaction(() => {
        for (const item of results) {
          const source = `${item.market}-reconciliation`;
          dependencies.repositories.sourceChecks.upsert({
            jobRunId: job.id, source, status: item.status,
            dataDate: item.result?.dataDate ?? null, recordCount: item.result?.events.length ?? 0,
            checkedAt, errorMessage: item.error ?? null,
          });
          if (!item.result || item.status === 'stale') continue;
          for (const record of item.result.events) {
            if (record.market !== item.market || !watched.has(`${record.market}:${record.stockCode}`)) continue;
            const company = dependencies.repositories.companies.findByStockCode(record.stockCode, record.market)!;
            const existing = dependencies.repositories.materialEvents.findBySourceKey(record.sourceKey);
            if (!existing && dependencies.repositories.materialEvents.findNearbyAnnouncement(company.id, record.title, record.publishedAt)?.source === 'mops') continue;
            const prepared = prepareMaterialEvent(record, existing ? [existing] : []);
            const event = dependencies.repositories.materialEvents.upsert({
              companyId: company.id,
              sourceKey: prepared.sourceKey, contentFingerprint: prepared.contentFingerprint,
              title: record.title, content: record.content, publishedAt: record.publishedAt,
              revisionOf: prepared.revisionOf, source: `${item.market.toLowerCase()}-reconciliation`,
              sourceUrl: record.sourceUrl, discoveredAt: checkedAt,
              eventType: /補充/.test(record.title) ? 'supplement' : /更正|修正/.test(record.title) ? 'correction' : 'announcement',
            });
            if (prepared.shouldNotify) discoveredIds.push(event.id);
          }
        }
        const candidates = discoveredIds.flatMap((eventId) => {
          const event = dependencies.repositories.materialEvents.find(eventId);
          const detail = event && dependencies.repositories.materialEvents.details(eventId);
          return event && detail ? [{ eventId, companyId: detail.companyId, companyName: detail.companyName, subject: event.title }] : [];
        });
        notification = buildMaterialNotification(candidates);
        if (notification) {
          dedupeKey = `material-reconciliation:${job.id}`;
          const intent = dependencies.repositories.notificationOutbox.enqueue({
            jobRunId: job.id, dedupeKey, channel: 'windows-toast', payloadJson: JSON.stringify(notification), createdAt: checkedAt,
          });
          deliverNotification = shouldDeliverNotification(intent);
        }
        dependencies.repositories.jobRuns.finish(job.id, {
          status, finishedAt: checkedAt, errorMessage: results.map(({ market, error }) => error && `${market}: ${error}`).filter(Boolean).join('; ') || undefined,
          summary: { markets: Object.fromEntries(results.map(({ market, result, status: sourceStatus }) => [market, { status: sourceStatus, dataDate: result?.dataDate ?? null }])), discovered: discoveredIds.length },
        });
        return { deliverNotification };
      });

      if (notification && dedupeKey && persisted.deliverNotification) {
        await deliverOutboxNotification({ repositories: dependencies.repositories, channel: dependencies.channel, now }, {
          jobRunId: job.id, dedupeKey, channel: 'windows-toast', payload: notification,
        });
      }
      return { jobRunId: job.id, status, discovered: discoveredIds.length };
    },
  };
}
