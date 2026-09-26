import { classifyDisclosureResult, type DisclosureStatus } from '../domain/disclosure-state';
import type { createRepositories } from '../repositories';
import type { DefaultDisclosureResult } from '../providers/default-disclosures';
import { buildDisclosureNotification } from './notification-messages';
import { deliverOutboxNotification, type NotificationChannel } from './notification-delivery';
import { shouldDeliverNotification } from './notification-outbox';

type Repositories = ReturnType<typeof createRepositories>;
type Provider = { fetchForDate(date: string): Promise<DefaultDisclosureResult> };
type Market = 'TWSE' | 'TPEX';
type MarketStatusSummary = { status: DisclosureStatus; dataDate: string | null; recordCount: number; retryAt: string | null; notifyEmpty?: boolean; errorMessage?: string };

function localTime(instant: string): string {
  const date = new Date(instant);
  if (Number.isNaN(date.valueOf())) throw new Error('違約交割檢查時間無效');
  return new Intl.DateTimeFormat('en-GB', { timeZone: 'Asia/Taipei', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' }).format(date);
}

function safeMessage(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  return message.replace(/\bBearer\s+[^\s,;]+/gi, 'Bearer [REDACTED]')
    .replace(/\b((?:access|refresh|oauth)?[_-]?token|client[_-]?secret|password|authorization(?:code)?)\s*[:=]\s*[^\s,;]+/gi, '$1=[REDACTED]');
}

export function createDefaultDisclosureMonitor(dependencies: {
  repositories: Repositories;
  providers: Record<Market, Provider>;
  channel: NotificationChannel;
  now?: () => string;
}) {
  const now = dependencies.now ?? (() => new Date().toISOString());
  return {
    async run(input: { targetDate: string; idempotencyKey: string }): Promise<{
      jobRunId: string;
      status: 'active' | 'complete' | 'degraded' | 'stale' | 'failed';
      markets: Record<Market, MarketStatusSummary>;
      notificationStatus: 'sent' | 'failed' | 'quiet';
      errorMessage?: string;
    }> {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(input.targetDate)) throw new Error('違約交割查詢日期必須使用 YYYY-MM-DD 格式');
      const startedAt = now();
      const job = dependencies.repositories.transaction(() => dependencies.repositories.jobRuns.start({
        kind: 'default-disclosure-monitoring', idempotencyKey: input.idempotencyKey, startedAt,
      }));
      if (!job.created) {
        const previous = JSON.parse(job.summaryJson) as { markets?: Record<Market, MarketStatusSummary> };
        const markets = previous.markets ?? {
          TWSE: { status: 'failed', dataDate: null, recordCount: 0, retryAt: null, errorMessage: '工作已建立但尚無完成摘要' },
          TPEX: { status: 'failed', dataDate: null, recordCount: 0, retryAt: null, errorMessage: '工作已建立但尚無完成摘要' },
        };
        return {
          jobRunId: job.id,
          status: job.status,
          markets: markets as Record<Market, MarketStatusSummary>,
          notificationStatus: 'quiet' as const,
          ...(job.errorMessage ? { errorMessage: job.errorMessage } : {}),
        };
      }
      const markets = ['TWSE', 'TPEX'] as const;
      const results = await Promise.all(markets.map(async (market) => {
        try {
          const result = await dependencies.providers[market].fetchForDate(input.targetDate);
          if (result.market !== market) throw new Error(`違約交割 provider 市場不符：${market}`);
          const state = classifyDisclosureResult({
            targetDate: input.targetDate, dataDate: result.dataDate, recordCount: result.records.length,
            checkedLocalTime: localTime(now()),
          });
          return { market, result, state, errorMessage: undefined as string | undefined };
        } catch (error) {
          const errorMessage = safeMessage(error);
          const state = classifyDisclosureResult({
            targetDate: input.targetDate, dataDate: null, recordCount: 0, checkedLocalTime: localTime(now()), error: errorMessage,
          });
          return { market, result: undefined, state, errorMessage };
        }
      }));

      const byMarket = new Map(results.map((item) => [item.market, item]));
      const states = { TWSE: byMarket.get('TWSE')!.state, TPEX: byMarket.get('TPEX')!.state };
      const failedCount = results.filter(({ state }) => state.status === 'failed').length;
      const staleCount = results.filter(({ state }) => state.status === 'stale').length;
      const hasSuccess = results.some(({ state }) => state.status === 'complete' || state.status === 'empty_success');
      const status = failedCount === results.length ? 'failed'
        : failedCount > 0 || (staleCount > 0 && hasSuccess) ? 'degraded'
          : staleCount > 0 ? 'stale' : 'complete';
      const errorMessage = results.filter(({ errorMessage }) => errorMessage).map(({ market, errorMessage: message }) => `${market}: ${message}`).join('; ') || undefined;
      const checkedAt = now();
      const persisted = dependencies.repositories.transaction(() => {
        for (const { market, result, state, errorMessage: marketError } of results) {
          const sourceStatus = state.status === 'empty_success' ? 'complete' : state.status;
          dependencies.repositories.sourceChecks.upsert({
            jobRunId: job.id, source: market, status: sourceStatus, dataDate: state.dataDate,
            recordCount: state.recordCount, checkedAt, errorMessage: marketError ?? null,
          });
          if (state.status !== 'complete' && state.status !== 'empty_success') continue;
          for (const disclosure of result?.records ?? []) {
            const company = dependencies.repositories.companies.findByStockCode(disclosure.stockCode, market);
            dependencies.repositories.defaultDisclosures.upsert({
              companyId: company?.id ?? null, market, disclosureDate: disclosure.disclosureDate,
              stockCode: disclosure.stockCode, brokerCode: disclosure.brokerCode, sourceKey: disclosure.sourceKey,
              disclosedAt: disclosure.disclosedAt, contentJson: JSON.stringify(disclosure.content),
            });
          }
        }
        const allRecords = results.flatMap(({ state, result }) =>
          state.status === 'complete' || state.status === 'empty_success' ? result?.records ?? [] : [],
        );
        const watchedCount = allRecords.filter(({ market, stockCode }) => dependencies.repositories.watchlist.isWatched(market, stockCode)).length;
        const watchedNames = allRecords.filter(({ market, stockCode }) => dependencies.repositories.watchlist.isWatched(market, stockCode))
          .map(({ companyName }) => companyName);
        const notifyEmpty = dependencies.repositories.settings.get<boolean>('notifyEmptyDefaultDisclosures') === true
          && results.every(({ state }) => state.status === 'empty_success');
        const notification = buildDisclosureNotification({ notifyEmpty, records: allRecords.map(({ companyName, market, stockCode }) => ({
          companyName: dependencies.repositories.watchlist.isWatched(market, stockCode) ? companyName : null,
        })) });
        const dedupeKey = notification ? `disclosure:${input.targetDate}:${job.id}` : undefined;
        const intent = notification && dedupeKey ? dependencies.repositories.notificationOutbox.enqueue({
          jobRunId: job.id, dedupeKey, channel: 'windows-toast', payloadJson: JSON.stringify(notification), createdAt: checkedAt,
        }) : undefined;
        dependencies.repositories.jobRuns.finish(job.id, {
          status, finishedAt: checkedAt,
          summary: { markets: states, recordCount: allRecords.length, watchedCount, watchedCompanies: watchedNames }, errorMessage,
        });
        return { notification, dedupeKey, deliverNotification: intent ? shouldDeliverNotification(intent) : false };
      });

      let notificationStatus: 'sent' | 'failed' | 'quiet' = 'quiet';
      if (persisted.notification && persisted.dedupeKey && persisted.deliverNotification) {
        notificationStatus = (await deliverOutboxNotification({ repositories: dependencies.repositories, channel: dependencies.channel, now }, {
          jobRunId: job.id, dedupeKey: persisted.dedupeKey, channel: 'windows-toast', payload: persisted.notification,
        })).status;
      }
      return {
        jobRunId: job.id, status,
        markets: states,
        notificationStatus, ...(errorMessage ? { errorMessage } : {}),
      };
    },
  };
}
