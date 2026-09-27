import { afterEach, describe, expect, it, vi } from 'vitest';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { createConferenceSyncService } from '../../src/services/conference-sync';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
afterEach(async () => { await paths?.cleanup(); paths = undefined; });

const conference = (CompId: string, CompName: string, Time: string) => ({
  CompId, CompName, Time, Location: '線上法人說明會', Content: 'fixture 內容',
  SourceUrl: `https://mops.twse.com.tw/mops/web/t100sb07_1?co_id=${CompId}`,
});

describe('conference-calendar-sync / service integration', () => {
  it('keeps a locally synced conference idempotent without querying Calendar again', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    let calendarCalls = 0;
    const service = createConferenceSyncService({
      repositories,
      calendar: {
        async hasEvent() { calendarCalls += 1; return false; },
        async insert() { calendarCalls += 1; return { id: 'existing-local-event-id' }; },
      },
      now: () => '2026-09-25T00:00:00.000Z',
    });
    try {
      const source = conference('2317', '測試電子股份有限公司', '115/09/29 時間：10 點 0 分 (24小時制)');
      const first = await service.sync([source]);
      expect(first[0]).toMatchObject({ status: 'synced' });
      expect(calendarCalls).toBe(2);

      const second = await service.sync([source]);
      expect(second[0]).toMatchObject({ status: 'synced', conferenceId: first[0].conferenceId });
      expect(calendarCalls).toBe(2);
      expect(repositories.calendarSyncs.find(first[0].conferenceId!)).toMatchObject({
        status: 'synced', calendarEventId: 'existing-local-event-id',
      });
    } finally { database.close(); }
  });

  it('keeps a previously found Calendar event marked existing on a repeated sync', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    let lookups = 0;
    let inserts = 0;
    const service = createConferenceSyncService({
      repositories,
      calendar: {
        async hasEvent() { lookups += 1; return lookups === 1; },
        async insert() { inserts += 1; return { id: 'should-not-exist' }; },
      },
      now: () => '2026-09-25T00:00:00.000Z',
    });
    try {
      const source = conference('2454', '聯發科技股份有限公司', '115/09/28 時間：14 點 30 分 (24小時制)');
      const first = await service.sync([source]);
      const second = await service.sync([source]);
      expect(first[0].status).toBe('existing');
      expect(second[0].status).toBe('existing');
      expect(lookups).toBe(1);
      expect(inserts).toBe(0);
    } finally { database.close(); }
  });

  it('stores sync outcomes, avoids existing calendar events, and continues after one invalid event', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const now = '2026-09-25T00:00:00.000Z';
    const calendarCalls: string[] = [];
    const inserts: string[] = [];
    const service = createConferenceSyncService({
      repositories,
      calendar: {
        async hasEvent(input) { calendarCalls.push(input.query); return input.query.startsWith('2454-'); },
        async insert(event) { inserts.push(event.summary); return { id: `calendar:${event.summary}` }; },
      },
      now: () => now,
    });
    try {
      const result = await service.sync([
        conference('2454', '聯發科技股份有限公司', '115/09/28 時間：14 點 30 分 (24小時制)'),
        conference('1101', '測試失敗股份有限公司', 'not a valid conference time'),
        conference('2317', '測試電子股份有限公司', '115/09/29 時間：10 點 0 分 (24小時制)'),
      ]);
      expect(result.map(({ status }) => status)).toEqual(['existing', 'failed', 'synced']);
      expect(calendarCalls).toEqual(['2454-聯發科技股份有限公司 法說會', '2317-測試電子股份有限公司 法說會']);
      expect(inserts).toEqual(['2317-測試電子股份有限公司 法說會']);
      expect(repositories.conferences.list()).toHaveLength(2);
      expect(repositories.conferences.find(result[0].conferenceId!)).toMatchObject({
        sourceUrl: 'https://mops.twse.com.tw/mops/web/t100sb07_1?co_id=2454',
      });
      expect(repositories.calendarSyncs.find(result[0].conferenceId!)).toMatchObject({ status: 'existing' });
      expect(repositories.calendarSyncs.find(result[2].conferenceId!)).toMatchObject({ status: 'synced', calendarEventId: 'calendar:2317-測試電子股份有限公司 法說會' });
    } finally { database.close(); }
  });

  it('does not query or create Calendar entries for expired conferences', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    let calls = 0;
    const service = createConferenceSyncService({
      repositories,
      calendar: { async hasEvent() { calls += 1; return false; }, async insert() { calls += 1; return {}; } },
      now: () => '2026-09-25T00:00:00.000Z',
    });
    try {
      const result = await service.sync([conference('2454', '聯發科技股份有限公司', '115/09/20 時間：14 點 30 分 (24小時制)')]);
      expect(result).toMatchObject([{ status: 'expired' }]);
      expect(calls).toBe(0);
    } finally { database.close(); }
  });

  it('records Calendar lookup failures, accepts nested event IDs, and keeps syncing the batch', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    repositories.companies.upsert({ market: 'TPEX', stockCode: '6488', name: '既有上櫃公司', updatedAt: '2026-09-25T00:00:00.000Z' });
    const service = createConferenceSyncService({
      repositories,
      calendar: {
        async hasEvent(input) {
          if (input.query.startsWith('2317-')) throw new Error('calendar temporarily unavailable');
          return false;
        },
        async insert(event) {
          if (event.summary.startsWith('6488-')) throw 'raw calendar insertion failure';
          return event.summary.startsWith('2308-') ? { data: { id: 'nested-calendar-id' } } : {};
        },
      },
      now: () => '2026-09-25T00:00:00.000Z',
    });
    try {
      const results = await service.sync([
        conference('2317', '查詢失敗公司', '115/09/29 時間：10 點 0 分 (24小時制)'),
        conference('2308', '巢狀事件代號公司', '115/09/29 時間：11 點 0 分 (24小時制)'),
        conference('2890', '無事件代號公司', '115/09/29 時間：12 點 0 分 (24小時制)'),
        conference('6488', '既有上櫃公司', '115/09/29 時間：13 點 0 分 (24小時制)'),
      ]);
      expect(results.map(({ status }) => status)).toEqual(['failed', 'synced', 'synced', 'failed']);
      expect(repositories.calendarSyncs.find(results[0].conferenceId!)).toMatchObject({ status: 'failed', lastError: 'calendar temporarily unavailable' });
      expect(repositories.calendarSyncs.find(results[1].conferenceId!)).toMatchObject({ status: 'synced', calendarEventId: 'nested-calendar-id' });
      expect(repositories.calendarSyncs.find(results[2].conferenceId!)).toMatchObject({ status: 'synced', calendarEventId: null });
      expect(results[3]).toMatchObject({ status: 'failed', errorMessage: 'raw calendar insertion failure' });
      expect(repositories.calendarSyncs.find(results[3].conferenceId!)).toMatchObject({ status: 'failed', lastError: 'raw calendar insertion failure' });
    } finally { database.close(); }
  });

  it('stops Calendar writes on rejected credentials and keeps all unsynced conferences pending', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    let calendarCalls = 0;
    const service = createConferenceSyncService({
      repositories,
      calendar: {
        async hasEvent() { calendarCalls += 1; const error = new Error('invalid_grant'); Object.assign(error, { code: 401 }); throw error; },
        async insert() { calendarCalls += 1; return {}; },
      },
      now: () => '2026-09-25T00:00:00.000Z',
    });
    try {
      const results = await service.sync([
        conference('2330', '測試公司', '115/09/29 時間：10 點 0 分 (24小時制)'),
        conference('2317', '另一測試公司', '115/09/29 時間：11 點 0 分 (24小時制)'),
      ]);
      expect(calendarCalls).toBe(1);
      expect(results.map(({ status }) => status)).toEqual(['failed', 'pending']);
      expect(repositories.conferences.list()).toHaveLength(2);
      expect(repositories.calendarSyncs.find(results[0].conferenceId!)).toMatchObject({ status: 'pending' });
      expect(repositories.calendarSyncs.find(results[1].conferenceId!)).toMatchObject({ status: 'pending' });
      expect(results[0]).toMatchObject({ authorizationRequired: true });
    } finally { database.close(); }
  });

  it('reports a storage failure per conference and does not call Calendar', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    repositories.transaction = () => { throw new Error('disk is read-only'); };
    let calendarCalls = 0;
    const service = createConferenceSyncService({
      repositories,
      calendar: { async hasEvent() { calendarCalls += 1; return false; }, async insert() { calendarCalls += 1; return {}; } },
      now: () => '2026-09-25T00:00:00.000Z',
    });
    try {
      const result = await service.sync([conference('2330', '測試公司', '115/09/29 時間：10 點 0 分 (24小時制)')]);
      expect(result).toMatchObject([{ status: 'failed', errorMessage: 'disk is read-only' }]);
      expect(calendarCalls).toBe(0);
      expect(repositories.conferences.list()).toHaveLength(0);
    } finally { database.close(); }
  });

  it('recognizes a string-form OAuth rejection and leaves the attempted conference pending', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const service = createConferenceSyncService({
      repositories,
      calendar: { async hasEvent() { throw 'invalid_grant'; }, async insert() { return {}; } },
      now: () => '2026-09-25T00:00:00.000Z',
    });
    try {
      const result = await service.sync([conference('2330', '測試公司', '115/09/29 時間：10 點 0 分 (24小時制)')]);
      expect(result).toMatchObject([{ status: 'failed', authorizationRequired: true }]);
      expect(repositories.calendarSyncs.find(result[0].conferenceId!)).toMatchObject({ status: 'pending' });
    } finally { database.close(); }
  });

  it('uses the injected fake system clock when no explicit clock port is supplied', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-09-25T00:00:00.000Z'));
    const service = createConferenceSyncService({
      repositories,
      calendar: { async hasEvent() { return false; }, async insert() { return { id: 'fake-clock-event' }; } },
    });
    try {
      const results = await service.sync([conference('2330', '測試公司', '115/09/29 時間：10 點 0 分 (24小時制)')]);
      expect(results).toMatchObject([{ status: 'synced' }]);
    } finally {
      vi.useRealTimers();
      database.close();
    }
  });
});
