import { afterEach, describe, expect, it } from 'vitest';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { createMaterialReconciliationMonitor } from '../../src/services/material-reconciliation-monitor';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
afterEach(async () => { await paths?.cleanup(); paths = undefined; });

describe('material-event-monitoring / reconciliation / stores a missed event from an official daily source', () => {
  it('does not duplicate an announcement already saved from the MOPS website while still backfilling a missing one', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const now = '2026-09-27T09:00:00.000Z';
    const company = repositories.companies.upsert({ market: 'TPEX', stockCode: '3629', name: '測試公司', updatedAt: now });
    repositories.watchlist.upsert({ companyId: company.id, active: true, category: '', notes: '', createdAt: now, updatedAt: now });
    repositories.materialEvents.upsert({
      companyId: company.id, sourceKey: 'mops:TPEX:3629:1150927:132950:1', contentFingerprint: 'old',
      title: '原公告', content: 'MOPS 完整內文', publishedAt: '2026-09-27T05:29:50.000Z', source: 'mops',
    });
    const base = {
      source: 'tpex' as const, market: 'TPEX' as const, stockCode: '3629', companyName: '測試公司',
      publishedAt: '2026-09-27T05:29:54.000Z', sourceUrl: 'https://mops.example/detail', revisionOf: null,
    };
    const monitor = createMaterialReconciliationMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { return { status: 'complete', dataDate: '2026-09-27', events: [] }; } },
        TPEX: { async fetchForDate() { return { status: 'complete', dataDate: '2026-09-27', events: [
          { ...base, sourceKey: 'tpex:same', title: '原公告', content: '對帳來源內文格式不同' },
          { ...base, sourceKey: 'tpex:missing', title: '漏接公告', content: '新公告完整內文' },
        ] }; } },
      },
      channel: { async send() {} }, now: () => now,
    });
    try {
      const result = await monitor.run({ targetDate: '2026-09-27', idempotencyKey: 'reconciliation:cross-source' });
      expect(result).toMatchObject({ status: 'complete', discovered: 1 });
      expect(repositories.materialEvents.list()).toHaveLength(2);
      expect(repositories.materialEvents.findBySourceKey('tpex:same')).toBeUndefined();
      expect(repositories.materialEvents.findBySourceKey('tpex:missing')).toMatchObject({ title: '漏接公告' });
    } finally { database.close(); }
  });

  it('persists only watched missing events with explicit reconciliation provenance and deduplicates the daily run', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const now = '2026-09-26T11:00:00.000Z';
    const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '台積電', updatedAt: now });
    repositories.watchlist.upsert({ companyId: company.id, active: true, category: '', notes: '', createdAt: now, updatedAt: now });
    let providerCalls = 0;
    const record = {
      source: 'twse' as const, market: 'TWSE' as const, stockCode: '2330', companyName: '台積電',
      publishedAt: '2026-09-26T09:00:00.000Z', title: '董事會決議', content: '完整對帳內容',
      sourceKey: 'twse:123', sourceUrl: 'https://mops.example/123', revisionOf: null,
    };
    const unmonitoredRecord = {
      ...record, market: 'TPEX' as const, source: 'tpex' as const, stockCode: '6488', companyName: '未關注公司',
      sourceKey: 'tpex:456',
    };
    const monitor = createMaterialReconciliationMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { providerCalls += 1; return { status: 'complete', dataDate: '2026-09-26', events: [record] }; } },
        TPEX: { async fetchForDate() { providerCalls += 1; return { status: 'complete', dataDate: '2026-09-26', events: [unmonitoredRecord] }; } },
      },
      channel: { async send() {} }, now: () => now,
    });

    try {
      const first = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'reconciliation:2026-09-26' });
      const repeated = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'reconciliation:2026-09-26' });
      const events = repositories.materialEvents.list();
      expect(first).toMatchObject({ status: 'complete', discovered: 1 });
      expect(repeated.jobRunId).toBe(first.jobRunId);
      expect(providerCalls).toBe(2);
      expect(events).toHaveLength(1);
      expect(events[0]).toMatchObject({ source: 'twse-reconciliation', title: '董事會決議', content: '完整對帳內容' });
    } finally { database.close(); }
  });

  it('keeps one market available and records a provider failure for the other', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const monitor = createMaterialReconciliationMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { throw new Error('TWSE timeout'); } },
        TPEX: { async fetchForDate() { return { status: 'complete', dataDate: '2026-09-26', events: [] }; } },
      },
      channel: { async send() {} }, now: () => '2026-09-26T11:00:00.000Z',
    });
    try {
      const result = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'reconciliation:partial' });
      expect(result.status).toBe('degraded');
      expect(repositories.sourceChecks.find(result.jobRunId, 'TWSE-reconciliation')).toMatchObject({ status: 'failed', errorMessage: 'TWSE timeout' });
      expect(repositories.sourceChecks.find(result.jobRunId, 'TPEX-reconciliation')).toMatchObject({ status: 'complete' });
    } finally { database.close(); }
  });

  it('records prior-day structured data as stale rather than an HTTP failure', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const monitor = createMaterialReconciliationMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { return { status: 'complete', dataDate: '2026-09-26', events: [] }; } },
        TPEX: { async fetchForDate() { return { status: 'complete', dataDate: '2026-09-25', events: [] }; } },
      },
      channel: { async send() {} }, now: () => '2026-09-26T11:00:00.000Z',
    });
    try {
      const partial = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'reconciliation:partial-stale' });
      expect(partial.status).toBe('degraded');
      expect(repositories.sourceChecks.find(partial.jobRunId, 'TPEX-reconciliation')).toMatchObject({
        status: 'stale', dataDate: '2026-09-25', errorMessage: null,
      });
      const staleMonitor = createMaterialReconciliationMonitor({
        repositories,
        providers: {
          TWSE: { async fetchForDate() { return { status: 'complete', dataDate: '2026-09-25', events: [] }; } },
          TPEX: { async fetchForDate() { return { status: 'complete', dataDate: '2026-09-25', events: [] }; } },
        },
        channel: { async send() {} }, now: () => '2026-09-26T11:00:00.000Z',
      });
      const stale = await staleMonitor.run({ targetDate: '2026-09-26', idempotencyKey: 'reconciliation:both-stale' });
      expect(stale.status).toBe('stale');
      expect(repositories.sourceChecks.find(stale.jobRunId, 'TWSE-reconciliation')).toMatchObject({ status: 'stale' });
    } finally { database.close(); }
  });

  it('returns the existing daily run without fetching providers a second time', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const now = '2026-09-26T11:00:00.000Z';
    const active = repositories.jobRuns.start({ kind: 'material-event-reconciliation', idempotencyKey: 'reconciliation:active', startedAt: now });
    let calls = 0;
    const monitor = createMaterialReconciliationMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { calls += 1; return { status: 'complete', dataDate: '2026-09-26', events: [] }; } },
        TPEX: { async fetchForDate() { calls += 1; return { status: 'complete', dataDate: '2026-09-26', events: [] }; } },
      },
      channel: { async send() {} }, now: () => now,
    });
    try {
      await expect(monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'reconciliation:active' })).resolves.toMatchObject({ jobRunId: active.id, status: 'degraded', discovered: 0 });
      expect(calls).toBe(0);
    } finally { database.close(); }
  });

  it('does not insert or notify reconciliation rows for inactive companies', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const inactive = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '已停用公司', updatedAt: '2026-09-26T00:00:00.000Z' });
    repositories.watchlist.upsert({ companyId: inactive.id, active: false, category: '', notes: '', createdAt: '2026-09-26T00:00:00.000Z', updatedAt: '2026-09-26T00:00:00.000Z' });
    const event = {
      source: 'twse' as const, market: 'TWSE' as const, stockCode: '2330', companyName: '已停用公司',
      publishedAt: '2026-09-26T09:00:00.000Z', title: '補充公告', content: '不保存', sourceKey: 'twse:inactive', sourceUrl: 'https://mops.example/evt', revisionOf: null,
    };
    const monitor = createMaterialReconciliationMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { return { status: 'complete', dataDate: '2026-09-26', events: [event] }; } },
        TPEX: { async fetchForDate() { return { status: 'complete', dataDate: '2026-09-26', events: [] }; } },
      },
      channel: { async send() { throw new Error('inactive companies must not notify'); } }, now: () => '2026-09-26T11:00:00.000Z',
    });
    try {
      const result = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'reconciliation:inactive' });
      expect(result).toMatchObject({ status: 'complete', discovered: 0 });
      expect(repositories.materialEvents.count()).toBe(0);
      expect(repositories.notificationOutbox.count()).toBe(0);
    } finally { database.close(); }
  });

  it('marks the job failed when both structured sources reject the daily request', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const monitor = createMaterialReconciliationMonitor({
      repositories,
      providers: {
        TWSE: { async fetchForDate() { throw 'TWSE offline'; } },
        TPEX: { async fetchForDate() { throw new Error('TPEX offline'); } },
      },
      channel: { async send() {} },
    });
    try {
      const result = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'reconciliation:failed' });
      expect(result.status).toBe('failed');
      expect(repositories.jobRuns.find(result.jobRunId)?.errorMessage).toContain('TWSE: TWSE offline');
      expect(repositories.sourceChecks.find(result.jobRunId, 'TPEX-reconciliation')).toMatchObject({ status: 'failed' });
    } finally { database.close(); }
  });
});
