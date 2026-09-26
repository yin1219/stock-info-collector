import Database, { type Database as SQLiteDatabase } from 'better-sqlite3';
import { existsSync } from 'node:fs';
import path from 'node:path';

export interface Migration {
  version: number;
  name?: string;
  up: (database: SQLiteDatabase) => void;
}

export interface OpenDatabaseOptions {
  migrations?: readonly Migration[];
  busyTimeoutMs?: number;
}

const initialMigrations: readonly Migration[] = [
  { version: 1, name: 'initial database', up: () => undefined },
  {
    version: 2,
    name: 'reporter business schema',
    up(database) {
      database.exec(`
        CREATE TABLE companies (
          id TEXT PRIMARY KEY NOT NULL,
          market TEXT NOT NULL CHECK (market IN ('TWSE', 'TPEX')),
          stock_code TEXT NOT NULL,
          name TEXT NOT NULL,
          updated_at TEXT NOT NULL,
          UNIQUE (market, stock_code)
        );

        CREATE TABLE watchlist_entries (
          id TEXT PRIMARY KEY NOT NULL,
          company_id TEXT NOT NULL UNIQUE REFERENCES companies(id) ON DELETE RESTRICT,
          active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0, 1)),
          category TEXT NOT NULL DEFAULT '',
          notes TEXT NOT NULL DEFAULT '',
          created_at TEXT NOT NULL,
          updated_at TEXT NOT NULL
        );

        CREATE TABLE material_events (
          id TEXT PRIMARY KEY NOT NULL,
          company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE RESTRICT,
          source_key TEXT NOT NULL UNIQUE,
          content_fingerprint TEXT NOT NULL,
          title TEXT NOT NULL,
          content TEXT NOT NULL,
          published_at TEXT NOT NULL,
          revision_of TEXT REFERENCES material_events(id) ON DELETE RESTRICT
        );

        CREATE TABLE default_disclosures (
          id TEXT PRIMARY KEY NOT NULL,
          company_id TEXT REFERENCES companies(id) ON DELETE RESTRICT,
          market TEXT NOT NULL CHECK (market IN ('TWSE', 'TPEX')),
          disclosure_date TEXT NOT NULL,
          stock_code TEXT NOT NULL,
          broker_code TEXT NOT NULL DEFAULT '',
          source_key TEXT NOT NULL UNIQUE,
          disclosed_at TEXT NOT NULL,
          content_json TEXT NOT NULL
        );

        CREATE TABLE conferences (
          id TEXT PRIMARY KEY NOT NULL,
          company_id TEXT NOT NULL REFERENCES companies(id) ON DELETE RESTRICT,
          stock_code TEXT NOT NULL,
          company_name TEXT NOT NULL,
          source_key TEXT NOT NULL UNIQUE,
          starts_at TEXT NOT NULL,
          location TEXT NOT NULL DEFAULT '',
          content TEXT NOT NULL DEFAULT ''
        );

        CREATE TABLE calendar_syncs (
          id TEXT PRIMARY KEY NOT NULL,
          conference_id TEXT NOT NULL UNIQUE REFERENCES conferences(id) ON DELETE RESTRICT,
          status TEXT NOT NULL CHECK (status IN ('pending', 'existing', 'synced', 'failed')),
          calendar_event_id TEXT,
          last_error TEXT,
          updated_at TEXT NOT NULL
        );

        CREATE TABLE job_runs (
          id TEXT PRIMARY KEY NOT NULL,
          kind TEXT NOT NULL,
          idempotency_key TEXT NOT NULL UNIQUE,
          status TEXT NOT NULL CHECK (status IN ('active', 'complete', 'degraded', 'stale', 'failed')),
          started_at TEXT NOT NULL,
          finished_at TEXT,
          summary_json TEXT NOT NULL DEFAULT '{}',
          error_message TEXT
        );

        CREATE TABLE source_checks (
          id TEXT PRIMARY KEY NOT NULL,
          job_run_id TEXT NOT NULL REFERENCES job_runs(id) ON DELETE CASCADE,
          source TEXT NOT NULL,
          status TEXT NOT NULL CHECK (status IN ('complete', 'degraded', 'stale', 'failed')),
          data_date TEXT,
          record_count INTEGER NOT NULL DEFAULT 0 CHECK (record_count >= 0),
          checked_at TEXT NOT NULL,
          error_message TEXT,
          UNIQUE (job_run_id, source)
        );

        CREATE TABLE notification_outbox (
          id TEXT PRIMARY KEY NOT NULL,
          job_run_id TEXT NOT NULL REFERENCES job_runs(id) ON DELETE CASCADE,
          event_id TEXT REFERENCES material_events(id) ON DELETE RESTRICT,
          dedupe_key TEXT NOT NULL UNIQUE,
          channel TEXT NOT NULL,
          payload_json TEXT NOT NULL,
          status TEXT NOT NULL CHECK (status IN ('pending', 'sent', 'failed')),
          created_at TEXT NOT NULL,
          sent_at TEXT
        );

        CREATE TABLE notification_deliveries (
          id TEXT PRIMARY KEY NOT NULL,
          outbox_id TEXT NOT NULL REFERENCES notification_outbox(id) ON DELETE CASCADE,
          attempt INTEGER NOT NULL CHECK (attempt > 0),
          status TEXT NOT NULL CHECK (status IN ('sent', 'failed')),
          attempted_at TEXT NOT NULL,
          error_message TEXT,
          UNIQUE (outbox_id, attempt)
        );

        CREATE TABLE settings (
          key TEXT PRIMARY KEY NOT NULL,
          value_json TEXT NOT NULL,
          updated_at TEXT NOT NULL
        );

        CREATE INDEX idx_material_events_published_at ON material_events(published_at);
        CREATE INDEX idx_default_disclosures_market_date ON default_disclosures(market, disclosure_date);
        CREATE INDEX idx_conferences_starts_at ON conferences(starts_at);
        CREATE INDEX idx_job_runs_started_at ON job_runs(started_at);
        CREATE INDEX idx_notification_outbox_status ON notification_outbox(status, created_at);
      `);
    },
  },
  {
    version: 3,
    name: 'material event provenance and read state',
    up(database) {
      database.exec(`
        ALTER TABLE material_events ADD COLUMN source TEXT NOT NULL DEFAULT 'mops';
        ALTER TABLE material_events ADD COLUMN source_url TEXT NOT NULL DEFAULT '';
        ALTER TABLE material_events ADD COLUMN discovered_at TEXT NOT NULL DEFAULT '';
        ALTER TABLE material_events ADD COLUMN read_at TEXT;
        ALTER TABLE material_events ADD COLUMN event_type TEXT NOT NULL DEFAULT 'announcement';
        CREATE INDEX idx_material_events_company_published ON material_events(company_id, published_at DESC);
      `);
    },
  },
];

export function openDatabase(filename: string, options: OpenDatabaseOptions = {}): SQLiteDatabase {
  const database = new Database(filename);
  const busyTimeoutMs = options.busyTimeoutMs ?? 5_000;
  if (!Number.isInteger(busyTimeoutMs) || busyTimeoutMs < 0) {
    database.close();
    throw new RangeError('SQLite busy timeout must be a non-negative integer');
  }

  try {
    database.pragma('foreign_keys = ON');
    database.pragma(`busy_timeout = ${busyTimeoutMs}`);
    database.pragma('journal_mode = WAL');
    runMigrations(database, options.migrations ?? initialMigrations);
    return database;
  } catch (error) {
    database.close();
    throw error;
  }
}

export function runMigrations(database: SQLiteDatabase, migrations: readonly Migration[]): void {
  database.exec(`
    CREATE TABLE IF NOT EXISTS schema_migrations (
      version INTEGER PRIMARY KEY NOT NULL,
      name TEXT NOT NULL,
      applied_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
    )
  `);

  const ordered = [...migrations].sort((left, right) => left.version - right.version);
  const seen = new Set<number>();
  for (const migration of ordered) {
    if (!Number.isSafeInteger(migration.version) || migration.version < 1 || seen.has(migration.version)) {
      throw new Error(`Invalid or duplicate migration version: ${migration.version}`);
    }
    seen.add(migration.version);
  }

  const currentVersion = database.prepare('SELECT MAX(version) AS version FROM schema_migrations')
    .get() as { version: number | null };
  const latestVersion = ordered.at(-1)?.version ?? 0;
  if ((currentVersion.version ?? 0) > latestVersion) {
    throw new Error('Database schema is newer than the application migration set');
  }

  let appliedVersion = currentVersion.version ?? 0;
  const recordMigration = database.prepare('INSERT INTO schema_migrations (version, name) VALUES (?, ?)');
  for (const migration of ordered) {
    if (migration.version <= appliedVersion) continue;
    if (appliedVersion > 0 && database.name !== ':memory:') {
      const backupPath = `${database.name}.pre-v${migration.version}.backup`;
      if (!existsSync(backupPath)) {
        database.prepare('VACUUM INTO ?').run(path.resolve(backupPath));
      }
    }
    const apply = database.transaction(() => {
      migration.up(database);
      recordMigration.run(migration.version, migration.name ?? `migration ${migration.version}`);
    });
    apply();
    appliedVersion = migration.version;
  }
}
