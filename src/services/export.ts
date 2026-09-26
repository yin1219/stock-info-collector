import { open, unlink } from 'node:fs/promises';
import type { Database as SQLiteDatabase } from 'better-sqlite3';

const dataTables = [
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
] as const;

function sanitizeSetting(value: unknown): { value?: unknown; containsSecret: boolean } {
  if (Array.isArray(value)) {
    const sanitized = value.map(sanitizeSetting);
    return { value: sanitized.map((entry) => entry.value), containsSecret: sanitized.some((entry) => entry.containsSecret) };
  }
  if (value === null || typeof value !== 'object') return { value, containsSecret: false };

  let containsSecret = false;
  const sanitized: Record<string, unknown> = {};
  for (const [key, nestedValue] of Object.entries(value)) {
    if (/token|secret|credential|password|authorization/i.test(key)) {
      containsSecret = true;
      continue;
    }
    const nested = sanitizeSetting(nestedValue);
    containsSecret ||= nested.containsSecret;
    sanitized[key] = nested.value;
  }
  return { value: sanitized, containsSecret };
}

export async function exportUserData(database: SQLiteDatabase, destination: string, exportedAt = new Date().toISOString()): Promise<string> {
  const exportData = database.transaction(() => {
    const tables = Object.fromEntries(dataTables.map((table) => [
      table,
      database.prepare(`SELECT * FROM ${table}`).all(),
    ]));
    const settings = database.prepare('SELECT key, value_json, updated_at AS updatedAt FROM settings').all() as Array<{
      key: string;
      value_json: string;
      updatedAt: string;
    }>;
    const safeSettings = settings.flatMap(({ key, value_json, updatedAt }) => {
      const { value, containsSecret } = sanitizeSetting(JSON.parse(value_json));
      return containsSecret ? [] : [{ key, value, updatedAt }];
    });

    return {
      format: 'stock-reporter-user-data',
      version: 1,
      exportedAt,
      data: tables,
      settings: safeSettings,
    };
  });

  const serialized = `${JSON.stringify(exportData(), null, 2)}\n`;
  const file = await open(destination, 'wx');
  let completed = false;
  try {
    await file.writeFile(serialized, 'utf8');
    await file.sync();
    completed = true;
    return destination;
  } finally {
    await file.close();
    if (!completed) await unlink(destination).catch(() => undefined);
  }
}
