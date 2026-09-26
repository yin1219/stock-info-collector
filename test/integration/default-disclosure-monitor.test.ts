import { afterEach, describe, expect, it } from 'vitest';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import type { DefaultDisclosureResult } from '../../src/providers/default-disclosures';
import { createDefaultDisclosureMonitor } from '../../src/services/default-disclosure-monitor';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
afterEach(async () => { await paths?.cleanup(); paths = undefined; });

async function setup() {
  paths = await createTestPaths();
  const database = openDatabase(paths.database);
  const repositories = createRepositories(database);
  const now = '2026-09-26T10:00:00.000Z';
  const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '關注半導體', updatedAt: now });
  repositories.watchlist.upsert({ companyId: company.id, active: true, category: '', notes: '', createdAt: now, updatedAt: now });
  return { database, repositories, now };
}

const result = (market: 'TWSE' | 'TPEX', code: string, dataDate = '2026-09-26'): DefaultDisclosureResult => ({
  market, dataDate, records: dataDate === '2026-09-26' ? [{
    market, disclosureDate: dataDate, stockCode: code, companyName: code === '2330' ? '關注半導體' : '一般公司',
    brokerCode: 'B001', disclosedAt: `${dataDate}T00:00:00+08:00`, sourceKey: `${market}:${code}:${dataDate}`,
    content: { amount: 50000, memo: 'fixture' },
  }] : [],
});

describe('default-disclosure-monitoring / cross-market persistence and job integration', () => {
  it('stores full-market rows and preserves the successful market when the other market fails', async () => {
    const { database, repositories } = await setup();
    const notifications: unknown[] = [];
    const monitor = createDefaultDisclosureMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { return result('TWSE', '2330'); } },
        TPEX: { async fetchForDate() { throw new Error('TPEX unavailable'); } },
      },
      channel: { async send(message) { notifications.push(message); } }, now: () => '2026-09-26T10:00:00.000Z',
    });
    try {
      const run = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'disclosure:run:1' });
      expect(run.status).toBe('degraded');
      expect(repositories.defaultDisclosures.list({ disclosureDate: '2026-09-26' })).toHaveLength(1);
      expect(repositories.sourceChecks.find(run.jobRunId, 'TWSE')).toMatchObject({ status: 'complete', recordCount: 1 });
      expect(repositories.sourceChecks.find(run.jobRunId, 'TPEX')).toMatchObject({ status: 'failed', errorMessage: 'TPEX unavailable' });
      expect(notifications[0]).toMatchObject({ body: expect.stringContaining('關注公司 1 筆') });
    } finally { database.close(); }
  });

  it('distinguishes stale official data from a valid empty current-day response and retries once at 19:00', async () => {
    const { database, repositories } = await setup();
    const monitor = createDefaultDisclosureMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { return result('TWSE', '2330', '2026-09-25'); } },
        TPEX: { async fetchForDate() { return { market: 'TPEX', dataDate: '2026-09-26', records: [] }; } },
      },
      channel: { async send() { throw new Error('empty success must not notify'); } }, now: () => '2026-09-26T10:30:00.000Z',
    });
    try {
      const run = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'disclosure:run:2' });
      expect(run.markets.TWSE).toMatchObject({ status: 'stale', retryAt: '2026-09-26T19:00:00+08:00' });
      expect(run.markets.TPEX).toMatchObject({ status: 'empty_success', recordCount: 0 });
      expect(run.notificationStatus).toBe('quiet');
    } finally { database.close(); }
  });

  it('does not refetch or redeliver a daily job whose idempotency key already completed', async () => {
    const { database, repositories } = await setup();
    let providerCalls = 0;
    let notifications = 0;
    const monitor = createDefaultDisclosureMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { providerCalls += 1; return result('TWSE', '2330'); } },
        TPEX: { async fetchForDate() { providerCalls += 1; return { market: 'TPEX', dataDate: '2026-09-26', records: [] }; } },
      },
      channel: { async send() { notifications += 1; } }, now: () => '2026-09-26T10:00:00.000Z',
    });
    try {
      const first = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'disclosure:once' });
      const repeated = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'disclosure:once' });
      expect(repeated.jobRunId).toBe(first.jobRunId);
      expect(repeated.notificationStatus).toBe('quiet');
      expect(providerCalls).toBe(2);
      expect(notifications).toBe(1);
    } finally { database.close(); }
  });

  it('notifies for current-day empty results only after the user explicitly enables the setting', async () => {
    const { database, repositories, now } = await setup();
    repositories.settings.set('notifyEmptyDefaultDisclosures', true, now);
    const messages: unknown[] = [];
    const monitor = createDefaultDisclosureMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { return { market: 'TWSE', dataDate: '2026-09-26', records: [] }; } },
        TPEX: { async fetchForDate() { return { market: 'TPEX', dataDate: '2026-09-26', records: [] }; } },
      },
      channel: { async send(message) { messages.push(message); } }, now: () => now,
    });
    try {
      const run = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'disclosure:empty-enabled' });
      expect(run.status).toBe('complete');
      expect(run.notificationStatus).toBe('sent');
      expect(messages).toMatchObject([{ body: '本日無違約交割揭露' }]);
    } finally { database.close(); }
  });

  it('marks wrong-market payloads as failed and keeps the total failed when every market errors', async () => {
    const { database, repositories } = await setup();
    const monitor = createDefaultDisclosureMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { return { ...result('TWSE', '2330'), market: 'TPEX' }; } },
        TPEX: { async fetchForDate() { throw 'TPEX request rejected'; } },
      },
      channel: { async send() { throw new Error('failed job must not notify'); } }, now: () => '2026-09-26T10:00:00.000Z',
    });
    try {
      const run = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'disclosure:both-failed' });
      expect(run.status).toBe('failed');
      expect(run.markets.TWSE.status).toBe('failed');
      expect(run.markets.TPEX.status).toBe('failed');
      expect(run.errorMessage).toContain('TPEX request rejected');
    } finally { database.close(); }
  });

  it('rejects malformed target dates before querying either market', async () => {
    const { database, repositories } = await setup();
    let calls = 0;
    const monitor = createDefaultDisclosureMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { calls += 1; return result('TWSE', '2330'); } },
        TPEX: { async fetchForDate() { calls += 1; return { market: 'TPEX', dataDate: '2026-09-26', records: [] }; } },
      },
      channel: { async send() {} },
    });
    try {
      await expect(monitor.run({ targetDate: 'tomorrow', idempotencyKey: 'disclosure:bad-date' })).rejects.toThrow('YYYY-MM-DD');
      expect(calls).toBe(0);
      expect(repositories.jobRuns.list('default-disclosure-monitoring')).toHaveLength(0);
    } finally { database.close(); }
  });
});
