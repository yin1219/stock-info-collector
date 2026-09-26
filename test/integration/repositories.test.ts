import { afterEach, describe, expect, it } from 'vitest';
import { createTestPaths } from '../support';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';

const cleanups: Array<() => Promise<void>> = [];

afterEach(async () => {
  await Promise.all(cleanups.splice(0).map((cleanup) => cleanup()));
});

async function openRepositories() {
  const paths = await createTestPaths();
  cleanups.push(paths.cleanup);
  const database = openDatabase(paths.database);
  return { database, repositories: createRepositories(database) };
}

describe('local-data-management / domain repositories', () => {
  it('lists material events with source details and persists read state without dropping history', async () => {
    const { database, repositories } = await openRepositories();
    const now = '2026-09-26T09:30:00.000Z';
    try {
      const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '台積電', updatedAt: now });
      const event = repositories.materialEvents.upsert({
        companyId: company.id, sourceKey: 'mops:event:1', contentFingerprint: 'hash', title: '董事會決議', content: '完整內容',
        publishedAt: now, source: 'mops', sourceUrl: 'https://mops.example/event/1', discoveredAt: now, eventType: 'announcement',
      });
      expect(event).toMatchObject({ source: 'mops', sourceUrl: 'https://mops.example/event/1', readAt: null, eventType: 'announcement' });
      expect(repositories.materialEvents.list({ unreadOnly: true })).toMatchObject([{ companyName: '台積電', stockCode: '2330', title: '董事會決議' }]);
      const other = repositories.materialEvents.upsert({
        companyId: company.id, sourceKey: 'mops:event:2', contentFingerprint: 'hash-2', title: '其他公告', content: '其他內容',
        publishedAt: '2026-09-26T09:31:00.000Z', source: 'mops', sourceUrl: 'https://mops.example/event/2', discoveredAt: now, eventType: 'announcement',
      });
      expect(repositories.materialEvents.list({ eventIds: [other.id] }).map(({ id }) => id)).toEqual([other.id]);
      expect(repositories.materialEvents.list({ eventIds: [] })).toEqual([]);
      repositories.materialEvents.markRead(event.id, '2026-09-26T10:00:00.000Z');
      repositories.materialEvents.markRead(other.id, '2026-09-26T10:00:00.000Z');
      expect(repositories.materialEvents.list({ unreadOnly: true })).toEqual([]);
      expect(repositories.materialEvents.find(event.id)?.readAt).toBe('2026-09-26T10:00:00.000Z');
      expect(repositories.materialEvents.findBySourceKey(event.sourceKey)?.id).toBe(event.id);
      expect(repositories.materialEvents.count()).toBe(2);
    } finally {
      database.close();
    }
  });

  it('transactionally upserts each domain record and returns the committed state', async () => {
    const { database, repositories } = await openRepositories();
    const now = '2026-09-26T00:00:00.000Z';
    try {
      const result = repositories.transaction(() => {
        const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '台積電', updatedAt: now });
        const companyAgain = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '台積電股份有限公司', updatedAt: now });
        expect(companyAgain.id).toBe(company.id);
        expect(repositories.companies.findByStockCode('2330', 'TWSE')?.name).toBe('台積電股份有限公司');

        const watch = repositories.watchlist.upsert({ companyId: company.id, active: true, category: '半導體', notes: '追蹤', createdAt: now, updatedAt: now });
        expect(repositories.watchlist.find(company.id)).toMatchObject({ id: watch.id, active: true, category: '半導體' });

        const event = repositories.materialEvents.upsert({ companyId: company.id, sourceKey: 'mops:evt:1', contentFingerprint: 'hash-1', title: '營收公告', content: '內容', publishedAt: now });
        expect(repositories.materialEvents.upsert({ companyId: company.id, sourceKey: 'mops:evt:1', contentFingerprint: 'hash-2', title: '更正公告', content: '更正內容', publishedAt: now }).id).toBe(event.id);
        expect(repositories.materialEvents.find(event.id)?.contentFingerprint).toBe('hash-2');

        const disclosure = repositories.defaultDisclosures.upsert({ companyId: company.id, market: 'TWSE', disclosureDate: '2026-09-25', stockCode: '2330', brokerCode: 'B001', sourceKey: 'twse:2026-09-25:2330:B001', disclosedAt: now, contentJson: '{}' });
        expect(repositories.defaultDisclosures.find(disclosure.id)?.brokerCode).toBe('B001');

        const conference = repositories.conferences.upsert({ companyId: company.id, stockCode: '2330', companyName: '台積電', sourceKey: 'mops:conference:1', startsAt: now, location: '台北', content: '法說會' });
        const sync = repositories.calendarSyncs.upsert({ conferenceId: conference.id, status: 'pending', updatedAt: now });
        expect(repositories.calendarSyncs.find(conference.id)?.id).toBe(sync.id);

        const job = repositories.jobRuns.start({ kind: 'material-event-monitoring', idempotencyKey: 'material:2026-09-26T00', startedAt: now });
        expect(repositories.jobRuns.start({ kind: 'material-event-monitoring', idempotencyKey: 'material:2026-09-26T00', startedAt: now })).toEqual({ ...job, created: false });
        const sourceCheck = repositories.sourceChecks.upsert({ jobRunId: job.id, source: 'MOPS', status: 'complete', dataDate: '2026-09-26', recordCount: 1, checkedAt: now });
        expect(repositories.sourceChecks.find(job.id, 'MOPS')?.id).toBe(sourceCheck.id);

        const notice = repositories.notificationOutbox.enqueue({ jobRunId: job.id, eventId: event.id, dedupeKey: 'notify:mops:evt:1', channel: 'windows-toast', payloadJson: '{"eventId":"evt-1"}', createdAt: now });
        expect(repositories.notificationOutbox.enqueue({ jobRunId: job.id, eventId: event.id, dedupeKey: 'notify:mops:evt:1', channel: 'windows-toast', payloadJson: '{}', createdAt: now }).id).toBe(notice.id);
        const delivery = repositories.notificationDeliveries.record({ outboxId: notice.id, attempt: 1, status: 'failed', attemptedAt: now, errorMessage: 'offline' });
        expect(repositories.notificationDeliveries.find(notice.id, 1)?.id).toBe(delivery.id);

        repositories.settings.set('monitoring', { enabled: true }, now);
        repositories.settings.set('monitoring', { enabled: false }, now);
        expect(repositories.settings.get<{ enabled: boolean }>('monitoring')).toEqual({ enabled: false });
        return company.id;
      });

      expect(result).toBeTruthy();
      expect(repositories.materialEvents.count()).toBe(1);
      expect(repositories.notificationOutbox.count()).toBe(1);
    } finally {
      database.close();
    }
  });

  it('rolls back all repository writes if a later write fails', async () => {
    const { database, repositories } = await openRepositories();
    const now = '2026-09-26T00:00:00.000Z';
    try {
      expect(() => repositories.transaction(() => {
        const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '1101', name: '台泥', updatedAt: now });
        repositories.materialEvents.upsert({ companyId: 'missing-company', sourceKey: 'invalid:evt', contentFingerprint: 'hash', title: 'bad', content: 'bad', publishedAt: now });
        return company.id;
      })).toThrow(/FOREIGN KEY constraint failed/);
      expect(repositories.companies.findByStockCode('1101')).toBeUndefined();
      expect(repositories.materialEvents.count()).toBe(0);
    } finally {
      database.close();
    }
  });
});
