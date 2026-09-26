import { afterEach, describe, expect, it } from 'vitest';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import type { MaterialProviderResult } from '../../src/providers/mops-material';
import { createMaterialMonitor } from '../../src/services/material-monitor';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
afterEach(async () => { await paths?.cleanup(); paths = undefined; });

async function setup() {
  paths = await createTestPaths();
  const database = openDatabase(paths.database);
  const repositories = createRepositories(database);
  const now = '2026-09-26T10:00:00.000Z';
  const watched = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '測試上市公司', updatedAt: now });
  const unwatched = repositories.companies.upsert({ market: 'TPEX', stockCode: '6488', name: '測試上櫃公司', updatedAt: now });
  repositories.watchlist.upsert({ companyId: watched.id, active: true, category: '', notes: '', createdAt: now, updatedAt: now });
  repositories.watchlist.upsert({ companyId: unwatched.id, active: false, category: '', notes: '', createdAt: now, updatedAt: now });
  return { database, repositories, now };
}

const event = (market: 'TWSE' | 'TPEX', stockCode: string, sourceKey: string) => ({
  source: 'mops' as const, market, stockCode, companyName: `公司${stockCode}`,
  publishedAt: '2026-09-26T09:30:00.000Z', title: `公告${stockCode}`, content: `完整公告內容${stockCode}`,
  sourceKey, sourceUrl: `https://mops.example/${sourceKey}`, revisionOf: null,
});

describe('material-event-monitoring / persistence and job integration', () => {
  it('persists source rows but queues and delivers only new events from active watched companies', async () => {
    const { database, repositories } = await setup();
    const deliveries: unknown[] = [];
    const result: MaterialProviderResult = { status: 'complete', dataDate: '2026-09-26', events: [event('TWSE', '2330', 'mops:1'), event('TPEX', '6488', 'mops:2')] };
    const monitor = createMaterialMonitor({
      repositories, primary: { async fetchForDate() { return result; } },
      channel: { async send(message) { deliveries.push(message); } }, now: () => '2026-09-26T10:00:00.000Z',
    });
    try {
      const run = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'material:run:1' });
      expect(run).toMatchObject({ status: 'complete', newEventIds: expect.arrayContaining([expect.any(String)]) });
      expect(repositories.materialEvents.count()).toBe(2);
      expect(repositories.notificationOutbox.count()).toBe(1);
      expect(deliveries).toHaveLength(1);
      expect(deliveries[0]).toMatchObject({ route: { type: 'event-detail' } });
      expect(repositories.sourceChecks.find(run.jobRunId, 'MOPS')).toMatchObject({ status: 'complete', recordCount: 2 });
      expect(repositories.jobRuns.find(run.jobRunId)).toMatchObject({ status: 'complete' });
    } finally { database.close(); }
  });

  it('does not notify an unchanged source item a second time', async () => {
    const { database, repositories } = await setup();
    let sends = 0;
    const record = event('TWSE', '2330', 'mops:repeat');
    const monitor = createMaterialMonitor({
      repositories, primary: { async fetchForDate(): Promise<MaterialProviderResult> { return { status: 'complete', dataDate: '2026-09-26', events: [record] }; } },
      channel: { async send() { sends += 1; } }, now: () => '2026-09-26T10:00:00.000Z',
    });
    try {
      await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'material:run:2' });
      await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'material:run:3' });
      expect(repositories.materialEvents.count()).toBe(1);
      expect(repositories.notificationOutbox.count()).toBe(1);
      expect(sends).toBe(1);
    } finally { database.close(); }
  });

  it('does not repeat a source check for an idempotency key that already completed', async () => {
    const { database, repositories } = await setup();
    let fetches = 0;
    const monitor = createMaterialMonitor({
      repositories,
      primary: { async fetchForDate(): Promise<MaterialProviderResult> { fetches += 1; return { status: 'complete', dataDate: '2026-09-26', events: [] }; } },
      channel: { async send() {} }, now: () => '2026-09-26T10:00:00.000Z',
    });
    try {
      const first = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'material:once' });
      const repeated = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'material:once' });
      expect(repeated.jobRunId).toBe(first.jobRunId);
      expect(fetches).toBe(1);
      expect(repeated.notificationStatus).toBe('quiet');
    } finally { database.close(); }
  });

  it('rejects malformed check dates before starting a durable job', async () => {
    const { database, repositories } = await setup();
    const monitor = createMaterialMonitor({
      repositories,
      primary: { async fetchForDate(): Promise<MaterialProviderResult> { throw new Error('must not be queried'); } },
      channel: { async send() {} },
    });
    try {
      await expect(monitor.run({ targetDate: 'tomorrow', idempotencyKey: 'material:invalid-date' }))
        .rejects.toThrow('YYYY-MM-DD');
      expect(repositories.jobRuns.list('material-event-monitoring')).toHaveLength(0);
    } finally { database.close(); }
  });

  it('keeps an unhealthy fallback stale and marks a primary degraded result as degraded', async () => {
    const { database, repositories } = await setup();
    const degraded = createMaterialMonitor({
      repositories,
      primary: { async fetchForDate(): Promise<MaterialProviderResult> { return { status: 'degraded', dataDate: '2026-09-26', events: [] }; } },
      fallback: { async fetchForDate() { throw 'rss unavailable'; } },
      channel: { async send() {} }, now: () => '2026-09-26T10:00:00.000Z',
    });
    const stale = createMaterialMonitor({
      repositories,
      primary: { async fetchForDate(): Promise<MaterialProviderResult> { return { status: 'complete', dataDate: '2026-09-25', events: [] }; } },
      fallback: { async fetchForDate() { throw new Error('rss still old'); } },
      channel: { async send() {} }, now: () => '2026-09-26T10:00:00.000Z',
    });
    try {
      const degradedRun = await degraded.run({ targetDate: '2026-09-26', idempotencyKey: 'material:primary-degraded' });
      const staleRun = await stale.run({ targetDate: '2026-09-26', idempotencyKey: 'material:stale-fallback-failed' });
      expect(degradedRun.status).toBe('degraded');
      expect(staleRun.status).toBe('stale');
      expect(repositories.sourceChecks.find(staleRun.jobRunId, 'MOPS-RSS')).toMatchObject({ status: 'failed', errorMessage: 'rss still old' });
    } finally { database.close(); }
  });

  it('stores and notifies a linked correction exactly once across repeated checks', async () => {
    const { database, repositories } = await setup();
    const deliveries: unknown[] = [];
    let current = event('TWSE', '2330', 'mops:revision');
    const monitor = createMaterialMonitor({
      repositories,
      primary: { async fetchForDate(): Promise<MaterialProviderResult> { return { status: 'complete', dataDate: '2026-09-26', events: [{ ...current }] }; } },
      channel: { async send(message) { deliveries.push(message); } }, now: () => '2026-09-26T10:00:00.000Z',
    });
    try {
      const original = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'material:revision:original' });
      current = { ...current, title: '董事會更正內容', content: '官方更正後完整內容' };
      const correction = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'material:revision:correction' });
      const repeated = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'material:revision:repeated' });
      const originalEvent = repositories.materialEvents.find(original.newEventIds[0]!);
      const correctedEvent = repositories.materialEvents.find(correction.newEventIds[0]!);
      expect(correctedEvent).toMatchObject({ revisionOf: originalEvent?.id, eventType: 'correction', content: '官方更正後完整內容' });
      expect(repeated.newEventIds).toEqual([]);
      expect(repositories.materialEvents.count()).toBe(2);
      expect(deliveries).toHaveLength(2);
    } finally { database.close(); }
  });

  it('reports degraded when the primary fails and fallback returns data', async () => {
    const { database, repositories } = await setup();
    const monitor = createMaterialMonitor({
      repositories,
      primary: { async fetchForDate() { throw new Error('primary offline'); } },
      fallback: { async fetchForDate(): Promise<MaterialProviderResult> { return { status: 'degraded', dataDate: '2026-09-26', events: [event('TWSE', '2330', 'rss:1')] }; } },
      channel: { async send() {} }, now: () => '2026-09-26T10:00:00.000Z',
    });
    try {
      const run = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'material:run:4' });
      expect(run).toMatchObject({ status: 'degraded' });
      expect(repositories.sourceChecks.find(run.jobRunId, 'MOPS')).toMatchObject({ status: 'failed' });
      expect(repositories.sourceChecks.find(run.jobRunId, 'MOPS-RSS')).toMatchObject({ status: 'degraded' });
    } finally { database.close(); }
  });

  it('distinguishes stale and failed results from an empty successful check', async () => {
    const { database, repositories } = await setup();
    const stale = createMaterialMonitor({
      repositories, primary: { async fetchForDate(): Promise<MaterialProviderResult> { return { status: 'complete', dataDate: '2026-09-25', events: [] }; } },
      channel: { async send() { throw new Error('should remain quiet'); } }, now: () => '2026-09-26T10:00:00.000Z',
    });
    const failed = createMaterialMonitor({
      repositories, primary: { async fetchForDate() { throw new Error('source unavailable'); } },
      channel: { async send() {} }, now: () => '2026-09-26T10:00:00.000Z',
    });
    try {
      const staleRun = await stale.run({ targetDate: '2026-09-26', idempotencyKey: 'material:run:5' });
      const failedRun = await failed.run({ targetDate: '2026-09-26', idempotencyKey: 'material:run:6' });
      expect(staleRun.status).toBe('stale');
      expect(failedRun.status).toBe('failed');
      expect(repositories.jobRuns.find(staleRun.jobRunId)?.status).toBe('stale');
      expect(repositories.jobRuns.find(failedRun.jobRunId)?.errorMessage).toContain('source unavailable');
      expect(repositories.notificationOutbox.count()).toBe(0);
    } finally { database.close(); }
  });
});
