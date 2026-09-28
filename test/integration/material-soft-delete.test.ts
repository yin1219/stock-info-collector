import { afterEach, describe, expect, it } from 'vitest';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { createMaterialMonitor } from '../../src/services/material-monitor';
import { createMaterialController } from '../../src/main/material-controller';
import { createMaterialReconciliationMonitor } from '../../src/services/material-reconciliation-monitor';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
afterEach(async () => { await paths?.cleanup(); paths = undefined; });

const record = {
  source: 'mops' as const, market: 'TWSE' as const, stockCode: '2330', companyName: '測試公司',
  sourceKey: 'mops:soft-delete:1', publishedAt: '2026-09-26T09:30:00.000Z',
  title: '測試公告', content: '完整公告', sourceUrl: 'https://mops.example/1', revisionOf: null,
};

describe('material-event-monitoring / soft deletion and notification retest', () => {
  it('hides one event without removing its row or prior notification, audits each delete and reacquisition', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const firstAt = '2026-09-26T10:00:00.000Z';
    const secondAt = '2026-09-26T10:10:00.000Z';
    const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '測試公司', updatedAt: firstAt });
    repositories.watchlist.upsert({ companyId: company.id, active: true, category: '', notes: '', createdAt: firstAt, updatedAt: firstAt });
    const event = repositories.materialEvents.upsert({ companyId: company.id, sourceKey: record.sourceKey, contentFingerprint: 'fingerprint', title: record.title, content: record.content, publishedAt: record.publishedAt });
    repositories.materialEvents.markRead(event.id, firstAt);
    const job = repositories.jobRuns.start({ kind: 'material-event-monitoring', idempotencyKey: 'prior-job', startedAt: firstAt });
    repositories.notificationOutbox.enqueue({ jobRunId: job.id, eventId: event.id, dedupeKey: 'prior-toast', channel: 'windows-toast', payloadJson: '{}', createdAt: firstAt });
    try {
      createMaterialController(repositories, () => secondAt).softDelete(event.id);
      expect(repositories.materialEvents.list()).toEqual([]);
      expect(repositories.materialEvents.count()).toBe(0);
      expect(repositories.materialEvents.find(event.id)).toMatchObject({ deletedAt: secondAt, readAt: firstAt });
      expect(repositories.notificationOutbox.count()).toBe(1);
      expect(createMaterialController(repositories, () => secondAt).detail(event.id)).toMatchObject({ id: event.id, deletedAt: secondAt });
      expect(repositories.materialEvents.deletionAudit(event.id)).toEqual([{ deletedAt: secondAt, reacquiredAt: null }]);
      expect(() => repositories.materialEvents.softDelete(event.id, secondAt)).toThrow(/已刪除/);
      const restored = repositories.materialEvents.upsert({ companyId: company.id, sourceKey: record.sourceKey, contentFingerprint: 'fingerprint', title: record.title, content: record.content, publishedAt: record.publishedAt, discoveredAt: '2026-09-26T10:20:00.000Z' });
      expect(restored).toMatchObject({ id: event.id, deletedAt: null, readAt: null });
      expect(repositories.materialEvents.list()).toHaveLength(1);
      expect(repositories.materialEvents.deletionAudit(event.id)).toEqual([{ deletedAt: secondAt, reacquiredAt: '2026-09-26T10:20:00.000Z' }]);
      expect(repositories.notificationOutbox.count()).toBe(1);
    } finally { database.close(); }
  });

  it('notifies again only after the same official item is reacquired, while retaining prior delivery history', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const at = '2026-09-26T10:00:00.000Z';
    const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '測試公司', updatedAt: at });
    repositories.watchlist.upsert({ companyId: company.id, active: true, category: '', notes: '', createdAt: at, updatedAt: at });
    let sends = 0;
    const monitor = createMaterialMonitor({ repositories, primary: { async fetchForDate() { return { status: 'complete' as const, dataDate: '2026-09-26', events: [record] }; } }, channel: { async send() { sends += 1; } }, now: () => at });
    try {
      const first = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'soft-delete:initial' });
      const id = first.newEventIds[0]!;
      expect(sends).toBe(1);
      repositories.materialEvents.softDelete(id, at);
      const second = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'soft-delete:retest' });
      expect(second.newEventIds).toEqual([id]);
      expect(sends).toBe(2);
      expect(repositories.notificationOutbox.count()).toBe(2);
      const third = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'soft-delete:unchanged' });
      expect(third.newEventIds).toEqual([]);
      expect(sends).toBe(2);
    } finally { database.close(); }
  });

  it('restores a deleted MOPS announcement from matching daily reconciliation without creating a duplicate', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const at = '2026-09-26T10:00:00.000Z';
    const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '測試公司', updatedAt: at });
    repositories.watchlist.upsert({ companyId: company.id, active: true, category: '', notes: '', createdAt: at, updatedAt: at });
    const original = repositories.materialEvents.upsert({ companyId: company.id, sourceKey: record.sourceKey, contentFingerprint: 'original', title: record.title, content: record.content, publishedAt: record.publishedAt, source: 'mops' });
    repositories.materialEvents.softDelete(original.id, at);
    let sends = 0;
    const monitor = createMaterialReconciliationMonitor({ repositories,
      providers: {
        TWSE: { async fetchForDate() { return { status: 'complete' as const, dataDate: '2026-09-26', events: [{ ...record, source: 'twse' as const, sourceKey: 'twse:recon:1', content: '對帳摘要' }] }; } },
        TPEX: { async fetchForDate() { return { status: 'complete' as const, dataDate: '2026-09-26', events: [] }; } },
      }, channel: { async send() { sends += 1; } }, now: () => at,
    });
    try {
      const result = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'soft-delete:cross-source' });
      expect(result.discovered).toBe(1);
      expect(repositories.materialEvents.list().map(({ id }) => id)).toEqual([original.id]);
      expect(repositories.materialEvents.findBySourceKey('twse:recon:1')?.id).toBe(original.id);
      expect(repositories.materialEvents.deletionAudit(original.id)).toEqual([{ deletedAt: at, reacquiredAt: at }]);
      expect(sends).toBe(1);
      const repeated = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'soft-delete:cross-source:again' });
      expect(repeated.discovered).toBe(0);
      expect(repositories.materialEvents.list()).toHaveLength(1);
      expect(sends).toBe(1);
    } finally { database.close(); }
  });

  it('does not duplicate a deleted reconciliation item after MOPS restores it with a different source key', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const at = '2026-09-26T10:00:00.000Z';
    const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '測試公司', updatedAt: at });
    repositories.watchlist.upsert({ companyId: company.id, active: true, category: '', notes: '', createdAt: at, updatedAt: at });
    const prior = repositories.materialEvents.upsert({ companyId: company.id, sourceKey: 'twse:recon:prior', contentFingerprint: 'old', title: record.title, content: '對帳摘要', publishedAt: record.publishedAt, source: 'twse-reconciliation' });
    repositories.materialEvents.softDelete(prior.id, at);
    let sends = 0;
    const monitor = createMaterialMonitor({ repositories, primary: { async fetchForDate() { return { status: 'complete' as const, dataDate: '2026-09-26', events: [record] }; } }, channel: { async send() { sends += 1; } }, now: () => at });
    try {
      const first = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'soft-delete:recon-to-mops:1' });
      const second = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'soft-delete:recon-to-mops:2' });
      expect(first.newEventIds).toEqual([prior.id]);
      expect(second.newEventIds).toEqual([]);
      expect(repositories.materialEvents.list().map(({ id }) => id)).toEqual([prior.id]);
      expect(sends).toBe(1);
      const reconciliation = createMaterialReconciliationMonitor({
        repositories,
        providers: {
          TWSE: { async fetchForDate() { return { status: 'complete' as const, dataDate: '2026-09-26', events: [{ ...record, source: 'twse' as const, sourceKey: prior.sourceKey, content: '對帳摘要' }] }; } },
          TPEX: { async fetchForDate() { return { status: 'complete' as const, dataDate: '2026-09-26', events: [] }; } },
        },
        channel: { async send() { sends += 1; } }, now: () => at,
      });
      const afterRestore = await reconciliation.run({ targetDate: '2026-09-26', idempotencyKey: 'soft-delete:recon-to-mops:reconcile' });
      expect(afterRestore.discovered).toBe(0);
      expect(repositories.materialEvents.list()).toHaveLength(1);
      expect(sends).toBe(1);
      repositories.materialEvents.softDelete(prior.id, at);
      const third = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'soft-delete:recon-to-mops:3' });
      const fourth = await monitor.run({ targetDate: '2026-09-26', idempotencyKey: 'soft-delete:recon-to-mops:4' });
      expect(third.newEventIds).toEqual([prior.id]);
      expect(fourth.newEventIds).toEqual([]);
      expect(repositories.materialEvents.list()).toHaveLength(1);
      expect(repositories.materialEvents.deletionAudit(prior.id)).toHaveLength(2);
      expect(sends).toBe(2);
    } finally { database.close(); }
  });
});
