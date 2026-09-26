import { afterEach, describe, expect, it } from 'vitest';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { deliverOutboxNotification } from '../../src/services/notification-delivery';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
afterEach(async () => {
  await paths?.cleanup();
  paths = undefined;
});

describe('notification-delivery / outbox / keeps event and failed delivery record', () => {
  it('persists notification intent before sending and records an OS channel failure', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const jobRun = repositories.jobRuns.start({ kind: 'material-monitor', idempotencyKey: 'run:2026-09-26:01', startedAt: '2026-09-26T00:00:00Z' });
    let outboxWasPersistedBeforeSend = false;

    const result = await deliverOutboxNotification({
      repositories,
      channel: {
        async send() {
          outboxWasPersistedBeforeSend = repositories.notificationOutbox.count() === 1;
          throw new Error('Windows notification unavailable');
        },
      },
      now: () => '2026-09-26T00:01:00Z',
    }, {
      jobRunId: jobRun.id,
      dedupeKey: 'run:2026-09-26:01:material-summary',
      channel: 'windows-toast',
      payload: { title: '重大訊息更新', body: '1 家公司、1 筆新事件', route: { type: 'event-list', eventIds: ['evt-1'] } },
    });

    expect(outboxWasPersistedBeforeSend).toBe(true);
    expect(result).toMatchObject({ status: 'failed', attempt: 1 });
    expect(repositories.notificationOutbox.count()).toBe(1);
    expect(database.prepare('SELECT status FROM notification_outbox').get()).toEqual({ status: 'failed' });
    expect(database.prepare('SELECT status, error_message AS errorMessage FROM notification_deliveries').get())
      .toEqual({ status: 'failed', errorMessage: 'Windows notification unavailable' });
    database.close();
  });

  it('does not redeliver an already sent outbox intent', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const job = repositories.jobRuns.start({ kind: 'material-monitor', idempotencyKey: 'run:sent', startedAt: '2026-09-26T00:00:00Z' });
    const item = repositories.notificationOutbox.enqueue({
      jobRunId: job.id, dedupeKey: 'sent-once', channel: 'windows-toast', payloadJson: '{}', createdAt: '2026-09-26T00:00:00Z',
    });
    repositories.notificationOutbox.setStatus(item.id, 'sent', '2026-09-26T00:01:00Z');
    let sends = 0;
    try {
      await expect(deliverOutboxNotification({
        repositories,
        channel: { async send() { sends += 1; } },
        now: () => '2026-09-26T00:02:00Z',
      }, {
        jobRunId: job.id, dedupeKey: 'sent-once', channel: 'windows-toast',
        payload: { title: '通知', body: '已送出', route: { type: 'disclosure-list' } },
      })).resolves.toEqual({ status: 'sent', attempt: 0 });
      expect(sends).toBe(0);
    } finally { database.close(); }
  });

  it('redacts bearer values and non-Error delivery failures in the attempt record', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const job = repositories.jobRuns.start({ kind: 'material-monitor', idempotencyKey: 'run:redact', startedAt: '2026-09-26T00:00:00Z' });
    try {
      const result = await deliverOutboxNotification({
        repositories, channel: { async send() { throw 'Bearer private-token'; } }, now: () => '2026-09-26T00:01:00Z',
      }, {
        jobRunId: job.id, dedupeKey: 'redacted-error', channel: 'windows-toast',
        payload: { title: '通知', body: '失敗', route: { type: 'disclosure-list' } },
      });
      expect(result.errorMessage).toBe('Bearer [REDACTED]');
      expect(database.prepare('SELECT error_message AS errorMessage FROM notification_deliveries').get())
        .toEqual({ errorMessage: 'Bearer [REDACTED]' });
    } finally { database.close(); }
  });
});
