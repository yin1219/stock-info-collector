import { afterEach, describe, expect, it, vi } from 'vitest';
import { createConferenceDailyJob } from '../../src/services/conference-daily-job';
import { createConferenceSyncService } from '../../src/services/conference-sync';
import { createGoogleCalendarController } from '../../src/main/google-calendar-controller';
import { createConferenceSyncGate } from '../../src/main/conference-sync-gate';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
afterEach(async () => { await paths?.cleanup(); paths = undefined; });

describe('conference-calendar-sync / scheduled entry / one daily owner and traceable outcome', () => {
  it('uses active SQLite codes and fake Calendar to trace expired, existing, new, and invalid rows without a window', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const now = '2026-09-25T00:00:00.000Z';
    const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2454', name: '聯發科技', updatedAt: now });
    repositories.watchlist.upsert({ companyId: company.id, active: true, category: '', notes: '', createdAt: now, updatedAt: now });
    const fetchConferences = vi.fn(async () => [
      { CompId: '2454', CompName: '已過期', Time: '115/09/20 時間：10 點 0 分 (24小時制)', Location: '線上', Content: '測試' },
      { CompId: '2454', CompName: '已存在', Time: '115/09/28 時間：10 點 0 分 (24小時制)', Location: '線上', Content: '測試' },
      { CompId: '2454', CompName: '新活動', Time: '115/09/29 時間：10 點 0 分 (24小時制)', Location: '線上', Content: '測試' },
      { CompId: '2454', CompName: '時間錯誤', Time: 'invalid', Location: '線上', Content: '測試' },
    ]);
    const insert = vi.fn(async () => ({ id: 'fake-new-event' }));
    const calendarSync = createConferenceSyncService({
      repositories, now: () => now,
      calendar: { hasEvent: async ({ query }) => query.includes('已存在'), insert },
    });
    const controller = createGoogleCalendarController({
      session: {
        async getStatus() { return { status: 'connected' }; },
        async connect() {}, async disconnect() {}, async markReauthorizationRequired() {},
      },
      repositories, fetchConferences,
      syncConferences: (rows) => calendarSync.sync(rows as Parameters<typeof calendarSync.sync>[0]),
      async selectCredentials() { return null; },
      async confirmLegacyTokenMigration() { return false; },
      async selectLegacyToken() { return null; },
    });
    const gate = createConferenceSyncGate({ taskOwner: async () => 'v2', runSync: () => controller.sync() });
    const job = createConferenceDailyJob({ jobRuns: repositories.jobRuns, taskOwner: async () => 'v2', sync: () => gate.sync(), now: () => now });
    try {
      expect(await job.run('2026-09-25')).toMatchObject({
        status: 'degraded', counts: { expired: 1, existing: 1, synced: 1, failed: 1 },
      });
      expect(fetchConferences).toHaveBeenCalledWith(['2454']);
      expect(insert).toHaveBeenCalledOnce();
      const summary = JSON.parse(repositories.jobRuns.list('conference-calendar-sync')[0].summaryJson) as {
        items: Array<{ status: string; summary: string }>;
      };
      expect(summary.items.map(({ status }) => status)).toEqual(['expired', 'existing', 'synced', 'failed']);
      expect(summary.items[3].summary).toContain('時間錯誤');
    } finally { database.close(); }
  });

  it('runs only a verified v2 task once per day and records the Calendar result', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const sync = vi.fn(async () => ({ status: 'complete', results: [{ status: 'synced' }, { status: 'existing' }] }));
    try {
      const job = createConferenceDailyJob({
        jobRuns: repositories.jobRuns, taskOwner: async () => 'v2', sync,
        now: () => '2026-09-29T11:00:00.000Z',
      });
      expect(await job.run('2026-09-29')).toMatchObject({ status: 'complete', counts: { synced: 1, existing: 1 } });
      expect(await job.run('2026-09-29')).toMatchObject({ status: 'already-checked' });
      expect(sync).toHaveBeenCalledTimes(1);
      expect(repositories.jobRuns.list('conference-calendar-sync')).toEqual([
        expect.objectContaining({ status: 'complete', idempotencyKey: 'conference-daily:2026-09-29' }),
      ]);
    } finally { database.close(); }
  });

  it('does not fetch or write Calendar when a v1 writer exists or task ownership is unknown', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const sync = vi.fn();
    try {
      for (const taskOwner of ['legacy', 'blocked'] as const) {
        const job = createConferenceDailyJob({ jobRuns: repositories.jobRuns, taskOwner: async () => taskOwner, sync });
        expect((await job.run('2026-09-29')).status).toBe('not-ready');
      }
      expect(sync).not.toHaveBeenCalled();
      expect(repositories.jobRuns.list('conference-calendar-sync')).toHaveLength(0);
    } finally { database.close(); }
  });

  it('rejects an overlapping run and stores a safe failure stage for a source reset', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    let rejectSource!: (error: Error) => void;
    const sync = vi.fn(() => new Promise<never>((_resolve, reject) => { rejectSource = reject; }));
    try {
      const job = createConferenceDailyJob({
        jobRuns: repositories.jobRuns, taskOwner: async () => 'v2', sync,
        now: () => '2026-09-29T11:00:00.000Z',
      });
      const first = job.run('2026-09-29');
      await vi.waitFor(() => expect(sync).toHaveBeenCalledOnce());
      expect((await job.run('2026-09-29')).status).toBe('already-running');
      rejectSource(new Error('法說會 2454 轉址查詢失敗：read ECONNRESET; access_token=secret'));
      expect(await first).toMatchObject({ status: 'failed', stage: 'MOPS 轉址' });
      const row = repositories.jobRuns.list('conference-calendar-sync')[0];
      expect(row.status).toBe('failed');
      expect(row.errorMessage).toContain('MOPS 轉址');
      expect(row.errorMessage).not.toContain('secret');
    } finally { database.close(); }
  });

  it('records partial rows and rejected Google authorization without calling them complete', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const sync = vi.fn()
      .mockResolvedValueOnce({ status: 'complete', results: [{ status: 'failed' }, { status: 'pending' }, { status: 'unknown' }] })
      .mockResolvedValueOnce({ status: 'reauthorization-required', results: [{ status: 'failed' }] });
    try {
      const job = createConferenceDailyJob({ jobRuns: repositories.jobRuns, taskOwner: async () => 'v2', sync });
      expect(await job.run('2026-09-29')).toMatchObject({ status: 'degraded', counts: { failed: 1, pending: 1 } });
      expect(await job.run('2026-09-30')).toMatchObject({ status: 'failed', counts: { failed: 1 } });
      expect(repositories.jobRuns.list('conference-calendar-sync').map(({ status }) => status)).toEqual(['failed', 'degraded']);
      expect(repositories.jobRuns.list('conference-calendar-sync')[0].errorMessage).toContain('授權未就緒');
      await expect(job.run('invalid-date')).rejects.toThrow(/日期格式/);
    } finally { database.close(); }
  });

  it('keeps a safe per-conference outcome in the daily job record after a partial failure', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const sync = vi.fn(async () => ({ status: 'complete', results: [
      { status: 'failed', summary: '2454-聯發科技 法說會', errorMessage: 'invalid date; access_token=secret' },
      { status: 'synced', summary: '2330-台積電 法說會', conferenceId: 'conference-1' },
    ] }));
    try {
      const job = createConferenceDailyJob({ jobRuns: repositories.jobRuns, taskOwner: async () => 'v2', sync });
      expect(await job.run('2026-09-29')).toMatchObject({ status: 'degraded', counts: { failed: 1, synced: 1 } });
      const recorded = JSON.parse(repositories.jobRuns.list('conference-calendar-sync')[0].summaryJson) as Record<string, unknown>;
      expect(recorded).toMatchObject({ failed: 1, synced: 1, items: [
        { status: 'failed', summary: '2454-聯發科技 法說會' },
        { status: 'synced', summary: '2330-台積電 法說會', conferenceId: 'conference-1' },
      ] });
      expect(JSON.stringify(recorded)).not.toContain('secret');
    } finally { database.close(); }
  });

  it('keeps source-page, Google, and unknown errors distinct without storing secret text', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const failures = [
      new Error('法說會 2454 結果頁讀取失敗：read ECONNRESET'),
      new Error('Google 授權失敗: refresh_token=private'),
      'unexpected raw failure',
    ];
    const onFailure = vi.fn();
    const sync = vi.fn(async () => { throw failures.shift(); });
    try {
      const job = createConferenceDailyJob({ jobRuns: repositories.jobRuns, taskOwner: async () => 'v2', sync, onFailure });
      expect(await job.run('2026-09-29')).toMatchObject({ status: 'failed', stage: 'MOPS 結果頁' });
      expect(await job.run('2026-09-30')).toMatchObject({ status: 'failed', stage: 'Google Calendar 授權' });
      expect(await job.run('2026-10-01')).toMatchObject({ status: 'failed', stage: '法說會同步' });
      expect(onFailure).toHaveBeenCalledTimes(3);
      expect(repositories.jobRuns.list('conference-calendar-sync').every(({ errorMessage }) => !errorMessage?.includes('private'))).toBe(true);
    } finally { database.close(); }
  });

  it('fails closed when task ownership inspection throws before creating a job row', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const sync = vi.fn();
    try {
      const job = createConferenceDailyJob({
        jobRuns: repositories.jobRuns,
        taskOwner: async () => { throw new Error('access denied'); }, sync,
      });
      expect(await job.run('2026-09-29')).toMatchObject({ status: 'not-ready' });
      expect(sync).not.toHaveBeenCalled();
      expect(repositories.jobRuns.list('conference-calendar-sync')).toHaveLength(0);
    } finally { database.close(); }
  });
});
