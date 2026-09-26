import { afterEach, describe, expect, it } from 'vitest';
import { createTestPaths } from '../support';
import { openDatabase } from '../../src/repositories/migrations';

const cleanups: Array<() => Promise<void>> = [];

afterEach(async () => {
  await Promise.all(cleanups.splice(0).map((cleanup) => cleanup()));
});

async function openTestDatabase() {
  const paths = await createTestPaths();
  cleanups.push(paths.cleanup);
  return openDatabase(paths.database);
}

function uniqueKeys(database: ReturnType<typeof openDatabase>, table: string): string[][] {
  const indexes = database.pragma(`index_list('${table}')`) as Array<{ name: string; unique: number }>;
  return indexes.filter((index) => index.unique === 1).map((index) => {
    const columns = database.pragma(`index_info('${index.name}')`) as Array<{ name: string }>;
    return columns.map((column) => column.name);
  });
}

describe('local-data-management / SQLite application schema', () => {
  it('stores material event provenance, read state, and version-link metadata', async () => {
    const database = await openTestDatabase();
    try {
      const columns = database.pragma('table_info(material_events)') as Array<{ name: string }>;
      expect(columns.map(({ name }) => name)).toEqual(expect.arrayContaining([
        'source', 'source_url', 'discovered_at', 'read_at', 'event_type',
      ]));
      expect(database.prepare('SELECT MAX(version) AS version FROM schema_migrations').get()).toEqual({ version: 3 });
    } finally {
      database.close();
    }
  });

  it('creates all durable business and delivery tables', async () => {
    const database = await openTestDatabase();
    try {
      const tables = database.prepare("SELECT name FROM sqlite_master WHERE type = 'table'")
        .all().map((row) => (row as { name: string }).name);
      expect(tables).toEqual(expect.arrayContaining([
        'companies',
        'watchlist_entries',
        'material_events',
        'default_disclosures',
        'conferences',
        'calendar_syncs',
        'job_runs',
        'source_checks',
        'notification_outbox',
        'notification_deliveries',
        'settings',
        'schema_migrations',
      ]));
    } finally {
      database.close();
    }
  });

  it('enforces stable and idempotent keys for persisted entities', async () => {
    const database = await openTestDatabase();
    try {
      const expectedUniqueKeys: Record<string, string[][]> = {
        companies: [['market', 'stock_code']],
        watchlist_entries: [['company_id']],
        material_events: [['source_key']],
        default_disclosures: [['source_key']],
        conferences: [['source_key']],
        calendar_syncs: [['conference_id']],
        job_runs: [['idempotency_key']],
        source_checks: [['job_run_id', 'source']],
        notification_outbox: [['dedupe_key']],
        notification_deliveries: [['outbox_id', 'attempt']],
        settings: [['key']],
      };
      for (const [table, keys] of Object.entries(expectedUniqueKeys)) {
        expect(uniqueKeys(database, table), table).toEqual(expect.arrayContaining(keys));
      }
    } finally {
      database.close();
    }
  });

  it('enforces ownership and history foreign keys', async () => {
    const database = await openTestDatabase();
    try {
      const companyForeignKey = database.pragma('foreign_key_list(watchlist_entries)')
        .map((key) => ({ table: key.table, from: key.from, to: key.to }));
      expect(companyForeignKey).toContainEqual({ table: 'companies', from: 'company_id', to: 'id' });
      expect(() => database.prepare(`
        INSERT INTO watchlist_entries (id, company_id, active, created_at, updated_at)
        VALUES ('watch-1', 'missing-company', 1, '2026-09-26T00:00:00.000Z', '2026-09-26T00:00:00.000Z')
      `).run()).toThrow(/FOREIGN KEY constraint failed/);
    } finally {
      database.close();
    }
  });
});
