import { materialContentFingerprint, prepareMaterialEvent } from '../domain/dedup';
import type { createRepositories } from '../repositories';
import type { MaterialProviderResult, MaterialSourceRecord } from '../providers/mops-material';
import { buildMaterialNotification } from './notification-messages';
import { deliverOutboxNotification, type NotificationChannel } from './notification-delivery';
import { shouldDeliverNotification } from './notification-outbox';
import { evaluateSourceHealth } from '../domain/source-health';
import { buildSourceFailureNotification } from './notification-messages';

type Repositories = ReturnType<typeof createRepositories>;
type Provider = { fetchForDate(date: string): Promise<MaterialProviderResult> };
type JobStatus = 'active' | 'complete' | 'degraded' | 'stale' | 'failed';

function messageOf(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  return message
    .replace(/\bBearer\s+[^\s,;]+/gi, 'Bearer [REDACTED]')
    .replace(/\b((?:access|refresh|oauth)?[_-]?token|client[_-]?secret|password|authorization(?:code)?)\s*[:=]\s*[^\s,;]+/gi, '$1=[REDACTED]');
}

function checkStatus(result: MaterialProviderResult, targetDate: string): 'complete' | 'degraded' | 'stale' | 'failed' {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(result.dataDate) || result.dataDate > targetDate) return 'failed';
  if (result.dataDate < targetDate) return 'stale';
  return result.status;
}

export function createMaterialMonitor(dependencies: {
  repositories: Repositories;
  primary: Provider;
  fallback?: Provider;
  channel: NotificationChannel;
  now?: () => string;
}) {
  const now = dependencies.now ?? (() => new Date().toISOString());

  return {
    async run(input: { targetDate: string; idempotencyKey: string }): Promise<{
      jobRunId: string;
      status: JobStatus;
      newEventIds: string[];
      notificationStatus: 'sent' | 'failed' | 'quiet';
      errorMessage?: string;
    }> {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(input.targetDate)) throw new Error('重大訊息查詢日期必須使用 YYYY-MM-DD 格式');
      const startedAt = now();
      const job = dependencies.repositories.transaction(() => dependencies.repositories.jobRuns.start({
        kind: 'material-event-monitoring', idempotencyKey: input.idempotencyKey, startedAt,
      }));
      if (!job.created) {
        return { jobRunId: job.id, status: job.status, newEventIds: [], notificationStatus: 'quiet', ...(job.errorMessage ? { errorMessage: job.errorMessage } : {}) };
      }

      let primaryResult: MaterialProviderResult | undefined;
      let primaryError: string | undefined;
      let primaryStatus: ReturnType<typeof checkStatus> = 'failed';
      try {
        primaryResult = await dependencies.primary.fetchForDate(input.targetDate);
        primaryStatus = checkStatus(primaryResult, input.targetDate);
      } catch (error) {
        primaryError = messageOf(error);
      }

      let fallbackResult: MaterialProviderResult | undefined;
      let fallbackError: string | undefined;
      let fallbackStatus: ReturnType<typeof checkStatus> | undefined;
      if ((!primaryResult || primaryStatus === 'stale' || primaryStatus === 'failed') && dependencies.fallback) {
        try {
          fallbackResult = await dependencies.fallback.fetchForDate(input.targetDate);
          fallbackStatus = checkStatus(fallbackResult, input.targetDate);
        } catch (error) {
          fallbackError = messageOf(error);
          fallbackStatus = 'failed';
        }
      }

      const fallbackUsable = fallbackResult && (fallbackStatus === 'complete' || fallbackStatus === 'degraded');
      const selected = fallbackUsable ? fallbackResult : primaryResult && (primaryStatus === 'complete' || primaryStatus === 'degraded') ? primaryResult : undefined;
      const finalStatus: JobStatus = fallbackUsable
        ? 'degraded'
        : selected ? (primaryStatus === 'degraded' ? 'degraded' : 'complete')
          : primaryStatus === 'stale' || fallbackStatus === 'stale' ? 'stale' : 'failed';
      const errorMessage = finalStatus === 'failed' ? [primaryError, fallbackError].filter(Boolean).join('; ') || '重大訊息來源回傳無效狀態' : undefined;

      const selectedSource = fallbackUsable ? 'MOPS-RSS' : 'MOPS';
      const records = selected?.events ?? [];
      const checkedAt = now();
      const sourceHealth = evaluateSourceHealth(dependencies.repositories.sourceChecks.history('MOPS'), primaryStatus);
      const sourceAlert = sourceHealth.notify ? buildSourceFailureNotification({
        source: 'MOPS', consecutiveFailures: sourceHealth.consecutiveFailures,
        errorMessage: primaryError ?? (primaryStatus === 'complete' ? null : `來源狀態：${primaryStatus}`),
      }) : null;
      const activeWatchlist = dependencies.repositories.watchlist.list({ activeOnly: true });
      const activeByKey = new Map(activeWatchlist.map((entry) => [`${entry.market}:${entry.stockCode}`, entry]));
      const persisted = dependencies.repositories.transaction(() => {
        dependencies.repositories.sourceChecks.upsert({
          jobRunId: job.id, source: 'MOPS', status: primaryStatus, dataDate: primaryResult?.dataDate ?? null,
          recordCount: primaryResult?.events.length ?? 0, checkedAt, errorMessage: primaryError ?? primaryResult?.warning ?? (primaryResult && primaryStatus === 'failed' ? 'MOPS 資料日期無效' : null),
        });
        if (dependencies.fallback && fallbackStatus) {
          dependencies.repositories.sourceChecks.upsert({
            jobRunId: job.id, source: 'MOPS-RSS', status: fallbackStatus, dataDate: fallbackResult?.dataDate ?? null,
            recordCount: fallbackResult?.events.length ?? 0, checkedAt, errorMessage: fallbackError ?? null,
          });
        }

        const newEventIds: string[] = [];
        const notifyCandidates: Array<{ eventId: string; companyId: string; companyName: string; subject: string }> = [];
        for (const record of records) {
          const company = dependencies.repositories.companies.upsert({
            market: record.market, stockCode: record.stockCode, name: record.companyName, updatedAt: checkedAt,
          });
          const nearby = dependencies.repositories.materialEvents.findNearbyAnnouncement(company.id, record.title, record.publishedAt);
          if (nearby?.source === `${record.market.toLowerCase()}-reconciliation` && !nearby.deletedAt) continue;
          const original = dependencies.repositories.materialEvents.findBySourceKey(record.sourceKey);
          if (original && original.sourceKey !== record.sourceKey && !original.deletedAt
            && original.contentFingerprint === materialContentFingerprint(record.content)) continue;
          const restoreFromAlias = original?.deletedAt && original.sourceKey !== record.sourceKey ? original : undefined;
          const restoreFromNearby = !original && nearby?.deletedAt ? nearby : undefined;
          const restoring = restoreFromAlias ?? restoreFromNearby;
          let prepared = prepareMaterialEvent(record, original && original.sourceKey === record.sourceKey ? [original] : []);
          if (prepared.kind === 'revision') {
            const revision = dependencies.repositories.materialEvents.findBySourceKey(prepared.sourceKey);
            if (revision && original) prepared = prepareMaterialEvent(record, [original, revision]);
          }
          const explicitKind = /補充/.test(record.title) ? 'supplement' : /更正|修正/.test(record.title) ? 'correction' : 'announcement';
          const event = dependencies.repositories.materialEvents.upsert({
            companyId: company.id, sourceKey: restoring?.sourceKey ?? prepared.sourceKey, contentFingerprint: prepared.contentFingerprint,
            title: record.title, content: record.content, publishedAt: record.publishedAt,
            revisionOf: restoring?.revisionOf ?? prepared.revisionOf, source: record.source, sourceUrl: record.sourceUrl,
            discoveredAt: checkedAt, eventType: prepared.kind === 'revision' ? explicitKind === 'announcement' ? 'correction' : explicitKind : explicitKind,
            ...(restoring && restoring.sourceKey !== record.sourceKey ? { reacquiredSourceKey: record.sourceKey } : {}),
          });
          if (prepared.shouldNotify || Boolean(original?.deletedAt) || Boolean(restoring)) {
            newEventIds.push(event.id);
            if (activeByKey.has(`${record.market}:${record.stockCode}`)) {
              notifyCandidates.push({ eventId: event.id, companyId: company.id, companyName: record.companyName, subject: record.title });
            }
          }
        }

        const notification = buildMaterialNotification(notifyCandidates);
        let notificationDedupeKey: string | undefined;
        let deliverNotification = false;
        if (notification) {
          notificationDedupeKey = `material:${job.id}:${notifyCandidates.map(({ eventId }) => eventId).join(',')}`;
          const intent = dependencies.repositories.notificationOutbox.enqueue({
            jobRunId: job.id, dedupeKey: notificationDedupeKey, channel: 'windows-toast',
            payloadJson: JSON.stringify(notification), createdAt: checkedAt,
            ...(notifyCandidates.length === 1 ? { eventId: notifyCandidates[0].eventId } : {}),
          });
          deliverNotification = shouldDeliverNotification(intent);
        }
        const sourceAlertDedupeKey = sourceAlert ? `source-health:MOPS:${job.id}` : undefined;
        const sourceAlertIntent = sourceAlert && sourceAlertDedupeKey ? dependencies.repositories.notificationOutbox.enqueue({
          jobRunId: job.id, dedupeKey: sourceAlertDedupeKey, channel: 'windows-toast',
          payloadJson: JSON.stringify(sourceAlert), createdAt: checkedAt,
        }) : undefined;
        dependencies.repositories.jobRuns.finish(job.id, {
          status: finalStatus, finishedAt: checkedAt,
          summary: { source: selectedSource, dataDate: selected?.dataDate ?? null, events: records.length, newEvents: newEventIds.length },
          errorMessage,
        });
        return { newEventIds, notification, notificationDedupeKey, deliverNotification, sourceAlertIntent, sourceAlertDedupeKey };
      });

      const deliveryStatuses: Array<'sent' | 'failed'> = [];
      if (persisted.notification && persisted.notificationDedupeKey && persisted.deliverNotification) {
        const delivery = await deliverOutboxNotification({
          repositories: dependencies.repositories,
          channel: dependencies.channel,
          now,
        }, {
          jobRunId: job.id, dedupeKey: persisted.notificationDedupeKey,
          channel: 'windows-toast', payload: persisted.notification,
        });
        deliveryStatuses.push(delivery.status);
      }
      if (sourceAlert && persisted.sourceAlertIntent && persisted.sourceAlertDedupeKey && shouldDeliverNotification(persisted.sourceAlertIntent)) {
        const delivery = await deliverOutboxNotification({ repositories: dependencies.repositories, channel: dependencies.channel, now }, {
          jobRunId: job.id, dedupeKey: persisted.sourceAlertDedupeKey, channel: 'windows-toast', payload: sourceAlert,
        });
        deliveryStatuses.push(delivery.status);
      }
      const notificationStatus: 'sent' | 'failed' | 'quiet' = deliveryStatuses.includes('sent') ? 'sent'
        : deliveryStatuses.length ? 'failed' : 'quiet';
      return { jobRunId: job.id, status: finalStatus, newEventIds: persisted.newEventIds, notificationStatus, ...(errorMessage ? { errorMessage } : {}) };
    },
  };
}
