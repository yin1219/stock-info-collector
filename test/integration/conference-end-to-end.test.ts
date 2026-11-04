import { afterEach, describe, expect, it, vi } from 'vitest';
import { createConferenceHttpClient } from '../../src/main/conference-http';
import { createGoogleCalendarController } from '../../src/main/google-calendar-controller';
import { createLegacyConferenceProvider } from '../../src/providers/conference';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { createConferenceSyncService } from '../../src/services/conference-sync';
import { createTestPaths, loadFixture, type TestPaths } from '../support';

let paths: TestPaths | undefined;
afterEach(async () => { await paths?.cleanup(); paths = undefined; });

describe('conference-calendar-sync / offline end-to-end pipeline', () => {
  it('recovers a reset MOPS redirect, saves a watched conference, and avoids a second Calendar insert', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const now = '2026-09-25T00:00:00.000Z';
    const watched = repositories.companies.upsert({ market: 'TWSE', stockCode: '2454', name: '聯發科技股份有限公司', updatedAt: now });
    repositories.watchlist.upsert({ companyId: watched.id, active: true, category: '', notes: '', createdAt: now, updatedAt: now });
    const sourceUrl = 'https://mopsov.twse.com.tw/mops/web/ajax_t100sb07_1?co_id=2454';
    const html = await loadFixture('conferences/mops-conference.html');
    const reset = Object.assign(new Error('read ECONNRESET'), { code: 'ECONNRESET' });
    const post = vi.fn().mockRejectedValueOnce(reset).mockResolvedValue({ data: { result: { url: sourceUrl } } });
    const get = vi.fn().mockResolvedValue({ data: html });
    const delay = vi.fn(async () => undefined);
    const provider = createLegacyConferenceProvider(createConferenceHttpClient({ post, get }, delay));
    let remoteExists = false;
    const hasEvent = vi.fn(async () => remoteExists);
    const insert = vi.fn(async () => { remoteExists = true; return { id: 'fake-calendar-event' }; });
    const sync = createConferenceSyncService({ repositories, calendar: { hasEvent, insert }, now: () => now });
    const controller = createGoogleCalendarController({
      session: {
        async getStatus() { return { status: 'connected' }; },
        async connect() {}, async disconnect() {}, async markReauthorizationRequired() {},
      },
      repositories,
      fetchConferences: (codes) => provider.fetch(codes),
      syncConferences: (conferences) => sync.sync(conferences as Parameters<typeof sync.sync>[0]),
      now: () => new Date(now),
      async selectCredentials() { return null; },
      async confirmLegacyTokenMigration() { return false; },
      async selectLegacyToken() { return null; },
    });

    try {
      const first = await controller.sync();
      expect(first).toMatchObject({ status: 'complete', results: [{ status: 'synced' }] });
      expect(post).toHaveBeenCalledTimes(2);
      expect(get).toHaveBeenCalledWith(sourceUrl, expect.any(Object));
      expect(delay).toHaveBeenCalledOnce();
      expect(repositories.conferences.list()).toEqual([expect.objectContaining({ stockCode: '2454', sourceUrl })]);
      expect(insert).toHaveBeenCalledOnce();

      const second = await controller.sync();
      expect(second).toMatchObject({ status: 'complete', results: [{ status: 'existing' }] });
      expect(hasEvent).toHaveBeenCalledTimes(2);
      expect(insert).toHaveBeenCalledOnce();
      expect(repositories.conferences.list()).toHaveLength(1);

      const inactive = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '非關注公司', updatedAt: now });
      repositories.conferences.upsert({
        companyId: inactive.id, stockCode: '2330', companyName: '非關注公司',
        sourceKey: 'inactive-future', startsAt: '2026-09-30T02:00:00.000Z',
        location: '線上', content: '不應補回',
      });
      repositories.conferences.upsert({
        companyId: watched.id, stockCode: '2454', companyName: '聯發科技股份有限公司',
        sourceKey: 'watched-expired', startsAt: '2026-09-20T02:00:00.000Z',
        location: '線上', content: '已過期',
      });
      vi.mocked(post).mockResolvedValueOnce({ data: { result: { url: sourceUrl } } });
      vi.mocked(get).mockResolvedValueOnce({ data: '<html><center>查無較新法說會</center></html>' });
      remoteExists = false;
      const third = await controller.sync();
      expect(third).toMatchObject({ status: 'complete', results: [{ status: 'synced' }] });
      expect(hasEvent).toHaveBeenCalledTimes(3);
      expect(insert).toHaveBeenCalledTimes(2);
      expect(insert.mock.calls[1][0]).toEqual(insert.mock.calls[0][0]);
      expect(repositories.conferences.list()).toHaveLength(3);
    } finally { database.close(); }
  });
});
