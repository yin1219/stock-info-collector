import Database from 'better-sqlite3';
import { unlink } from 'node:fs/promises';
import { afterEach, describe, expect, it } from 'vitest';
import { createTestPaths } from '../support';
import { openDatabase, runMigrations, type Migration } from '../../src/repositories/migrations';

const cleanups: Array<() => Promise<void>> = [];

afterEach(async () => {
  await Promise.all(cleanups.splice(0).map((cleanup) => cleanup()));
});

async function createDatabasePath(): Promise<string> {
  const paths = await createTestPaths();
  cleanups.push(paths.cleanup);
  return paths.database;
}

describe('local-data-management / SQLite migrations', () => {
  it('configures WAL, foreign keys and a bounded busy timeout on a new database', async () => {
    const database = openDatabase(await createDatabasePath());
    try {
      expect(database.pragma('journal_mode', { simple: true })).toBe('wal');
      expect(database.pragma('foreign_keys', { simple: true })).toBe(1);
      expect(database.pragma('busy_timeout', { simple: true })).toBe(5_000);
    } finally {
      database.close();
    }
  });

  it('records migration versions and makes reopening idempotent', async () => {
    const databasePath = await createDatabasePath();
    const first = openDatabase(databasePath);
    first.close();

    const reopened = openDatabase(databasePath);
    try {
      expect(reopened.prepare('SELECT version FROM schema_migrations ORDER BY version').all())
        .toEqual([{ version: 1 }, { version: 2 }, { version: 3 }, { version: 4 }, { version: 5 }]);
    } finally {
      reopened.close();
    }
  });

  it('backs up a v4 database before adding soft deletion and keeps existing announcements', async () => {
    const databasePath = await createDatabasePath();
    const prior = openDatabase(databasePath);
    try {
      prior.prepare('INSERT INTO companies (id, market, stock_code, name, updated_at) VALUES (?, ?, ?, ?, ?)')
        .run('company-1', 'TWSE', '2330', '測試公司', '2026-09-26T10:00:00Z');
      prior.prepare('INSERT INTO material_events (id, company_id, source_key, content_fingerprint, title, content, published_at) VALUES (?, ?, ?, ?, ?, ?, ?)')
        .run('event-1', 'company-1', 'mops:old', 'hash', '舊公告', '完整內文', '2026-09-26T10:00:00Z');
      prior.exec('DROP TABLE material_event_deletion_audit; ALTER TABLE material_events DROP COLUMN deleted_at; DELETE FROM schema_migrations WHERE version = 5');
    } finally { prior.close(); }

    await unlink(`${databasePath}.pre-v5.backup`);

    const upgraded = openDatabase(databasePath);
    try {
      expect(upgraded.prepare('SELECT deleted_at AS deletedAt FROM material_events WHERE id = ?').get('event-1'))
        .toEqual({ deletedAt: null });
      expect(upgraded.prepare('SELECT MAX(version) AS version FROM schema_migrations').get()).toEqual({ version: 5 });
    } finally { upgraded.close(); }

    const backup = new Database(`${databasePath}.pre-v5.backup`, { readonly: true });
    try {
      expect(backup.prepare('SELECT MAX(version) AS version FROM schema_migrations').get()).toEqual({ version: 4 });
      expect(backup.prepare('SELECT content FROM material_events WHERE id = ?').get('event-1'))
        .toEqual({ content: '完整內文' });
    } finally { backup.close(); }
  });

  it('rolls back only the failing migration and rejects startup', async () => {
    const databasePath = await createDatabasePath();
    const migrations: Migration[] = [
      {
        version: 1,
        up(database) {
          database.exec('CREATE TABLE durable (value TEXT NOT NULL)');
          database.prepare('INSERT INTO durable (value) VALUES (?)').run('before-upgrade');
        },
      },
      {
        version: 2,
        up(database) {
          database.exec('CREATE TABLE partial (value TEXT NOT NULL)');
          throw new Error('injected migration failure');
        },
      },
    ];

    expect(() => openDatabase(databasePath, { migrations })).toThrow('injected migration failure');

    const backup = new Database(`${databasePath}.pre-v2.backup`);
    try {
      expect(backup.prepare('SELECT value FROM durable').all()).toEqual([{ value: 'before-upgrade' }]);
      expect(backup.prepare('SELECT version FROM schema_migrations ORDER BY version').all()).toEqual([{ version: 1 }]);
    } finally {
      backup.close();
    }

    expect(() => openDatabase(databasePath, { migrations })).toThrow('injected migration failure');

    const inspect = new Database(databasePath);
    try {
      expect(inspect.prepare('SELECT value FROM durable').all()).toEqual([{ value: 'before-upgrade' }]);
      expect(inspect.prepare('SELECT version FROM schema_migrations ORDER BY version').all())
        .toEqual([{ version: 1 }]);
      expect(inspect.prepare("SELECT name FROM sqlite_master WHERE name = 'partial'").all()).toEqual([]);
    } finally {
      inspect.close();
    }
  });

  it('rejects invalid timeouts, duplicate migration versions, and a database newer than the app', async () => {
    const databasePath = await createDatabasePath();
    expect(() => openDatabase(databasePath, { busyTimeoutMs: -1 })).toThrow(RangeError);

    const emptyDatabase = new Database(':memory:');
    try { runMigrations(emptyDatabase, []); } finally { emptyDatabase.close(); }

    const invalidVersions = new Database(':memory:');
    try {
      expect(() => runMigrations(invalidVersions, [{ version: 0, up() {} }])).toThrow('Invalid or duplicate migration version');
      expect(() => runMigrations(invalidVersions, [{ version: 1, up() {} }, { version: 1, up() {} }]))
        .toThrow('Invalid or duplicate migration version');
      invalidVersions.prepare('INSERT INTO schema_migrations (version, name) VALUES (?, ?)').run(3, 'future');
      expect(() => runMigrations(invalidVersions, [{ version: 1, up() {} }, { version: 2, up() {} }]))
        .toThrow('Database schema is newer than the application migration set');
    } finally {
      invalidVersions.close();
    }
  });
});
