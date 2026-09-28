import { afterEach, describe, expect, it } from 'vitest';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { createOutboxRecovery } from '../../src/services/notification-recovery';
import { createTestPaths, type TestPaths } from '../support';

let paths: TestPaths | undefined;
afterEach(async () => { await paths?.cleanup(); paths = undefined; });

describe('notification-delivery / recovery after abnormal restart', () => {
  it('delivers a committed pending intent once after reopening SQLite', async () => {
    paths = await createTestPaths();
    const firstDatabase = openDatabase(paths.database);
    const first = createRepositories(firstDatabase);
    const startedAt = '2026-09-26T10:00:00.000Z';
    const job = first.jobRuns.start({ kind: 'material-event-monitoring', idempotencyKey: 'recovery:pending', startedAt });
    first.notificationOutbox.enqueue({
      jobRunId: job.id, dedupeKey: 'recovery:pending', channel: 'windows-toast', createdAt: startedAt,
      payloadJson: JSON.stringify({ title: '測試公告', body: '待送通知', route: { type: 'source-status' } }),
    });
    firstDatabase.close();

    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const sent: unknown[] = [];
    const recovery = createOutboxRecovery({ repositories, channel: { async send(message) { sent.push(message); } }, now: () => '2026-09-26T10:05:00.000Z' });
    try {
      expect(await recovery.drain()).toEqual({ attempted: 1, sent: 1, failed: 0 });
      expect(await recovery.drain()).toEqual({ attempted: 0, sent: 0, failed: 0 });
      expect(sent).toHaveLength(1);
      expect(database.prepare('SELECT status FROM notification_outbox WHERE dedupe_key = ?').get('recovery:pending')).toEqual({ status: 'sent' });
      expect(database.prepare('SELECT attempt, status FROM notification_deliveries').all()).toEqual([{ attempt: 1, status: 'sent' }]);
    } finally { database.close(); }
  });

  it('retries a failed Windows delivery after a cooldown without resending after success', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const job = repositories.jobRuns.start({ kind: 'material-event-monitoring', idempotencyKey: 'recovery:failed', startedAt: '2026-09-26T10:00:00.000Z' });
    repositories.notificationOutbox.enqueue({
      jobRunId: job.id, dedupeKey: 'recovery:failed', channel: 'windows-toast', createdAt: '2026-09-26T10:00:00.000Z',
      payloadJson: JSON.stringify({ title: '測試公告', body: '重試通知', route: { type: 'source-status' } }),
    });
    let clock = '2026-09-26T10:05:00.000Z';
    let sends = 0;
    const recovery = createOutboxRecovery({
      repositories, now: () => clock,
      channel: { async send() { sends += 1; if (sends === 1) throw new Error('Windows 拒絕通知'); } },
    });
    try {
      expect(await recovery.drain()).toEqual({ attempted: 1, sent: 0, failed: 1 });
      expect(await recovery.drain()).toEqual({ attempted: 0, sent: 0, failed: 0 });
      clock = '2026-09-26T10:07:00.000Z';
      expect(await recovery.drain()).toEqual({ attempted: 1, sent: 1, failed: 0 });
      expect(await recovery.drain()).toEqual({ attempted: 0, sent: 0, failed: 0 });
      expect(sends).toBe(2);
      expect(database.prepare('SELECT attempt, status FROM notification_deliveries ORDER BY attempt').all())
        .toEqual([{ attempt: 1, status: 'failed' }, { attempt: 2, status: 'sent' }]);
    } finally { database.close(); }
  });

  it('uses the local clock when none is injected and returns without work for an empty outbox', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    try {
      const recovery = createOutboxRecovery({ repositories: createRepositories(database), channel: { async send() { throw new Error('unexpected'); } } });
      expect(await recovery.drain()).toEqual({ attempted: 0, sent: 0, failed: 0 });
    } finally { database.close(); }
  });

  it('does not start a second drain while a delivery is in progress', async () => {
    paths = await createTestPaths();
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const job = repositories.jobRuns.start({ kind: 'material-event-monitoring', idempotencyKey: 'recovery:overlap', startedAt: '2026-09-26T10:00:00.000Z' });
    repositories.notificationOutbox.enqueue({
      jobRunId: job.id, dedupeKey: 'recovery:overlap', channel: 'windows-toast', createdAt: '2026-09-26T10:00:00.000Z',
      payloadJson: JSON.stringify({ title: '測試公告', body: '待送通知', route: { type: 'source-status' } }),
    });
    let release!: () => void;
    let started!: () => void;
    const deliveryStarted = new Promise<void>((resolve) => { started = resolve; });
    const waitForRelease = new Promise<void>((resolve) => { release = resolve; });
    let sends = 0;
    const recovery = createOutboxRecovery({ repositories, now: () => '2026-09-26T10:05:00.000Z',
      channel: { async send() { sends += 1; started(); await waitForRelease; } } });
    try {
      const first = recovery.drain();
      await deliveryStarted;
      expect(await recovery.drain()).toEqual({ attempted: 0, sent: 0, failed: 0 });
      release();
      expect(await first).toEqual({ attempted: 1, sent: 1, failed: 0 });
      expect(sends).toBe(1);
    } finally { release(); database.close(); }
  });
});
