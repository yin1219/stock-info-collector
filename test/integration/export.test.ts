import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { createTestPaths } from '../support';
import { openDatabase } from '../../src/repositories/migrations';
import { createRepositories } from '../../src/repositories';
import { exportUserData } from '../../src/services/export';

const cleanups: Array<() => Promise<void>> = [];

afterEach(async () => {
  await Promise.all(cleanups.splice(0).map((cleanup) => cleanup()));
});

describe('local-data-management / versioned export', () => {
  it('exports business data and settings in a readable versioned file without reusable tokens', async () => {
    const paths = await createTestPaths();
    cleanups.push(paths.cleanup);
    const database = openDatabase(paths.database);
    const repositories = createRepositories(database);
    const now = '2026-09-26T00:00:00.000Z';
    const company = repositories.companies.upsert({ market: 'TWSE', stockCode: '2330', name: '台積電', updatedAt: now });
    const event = repositories.materialEvents.upsert({ companyId: company.id, sourceKey: 'mops:export:1', contentFingerprint: 'hash', title: '測試公告', content: '內容', publishedAt: now });
    repositories.materialEvents.softDelete(event.id, now);
    repositories.settings.set('monitoring', { enabled: true }, now);
    repositories.settings.set('safe-list', ['新聞', 3, null], now);
    repositories.settings.set('nested-secret-list', ['安全', { token: 'ARRAY-TOKEN-MUST-NOT-EXPORT' }], now);
    repositories.settings.set('test-only-sensitive-value', {
      access_token: 'ACCESS-TOKEN-MUST-NOT-EXPORT',
      nested: { refreshToken: 'REFRESH-TOKEN-MUST-NOT-EXPORT' },
      clientSecret: 'CLIENT-SECRET-MUST-NOT-EXPORT',
    }, now);

    try {
      const destination = path.join(paths.root, 'reporter-export.json');
      await exportUserData(database, destination, now);
      const serialized = await readFile(destination, 'utf8');
      const exported = JSON.parse(serialized) as {
        format: string;
        version: number;
        settings: Array<{ key: string; value: unknown }>;
        data: { companies: Array<{ stock_code: string; name: string }>; material_events: Array<{ id: string; deleted_at: string | null }>; material_event_deletion_audit: Array<{ event_id: string; deleted_at: string }> };
      };

      expect(exported.format).toBe('stock-reporter-user-data');
      expect(exported.version).toBe(1);
      expect(exported.data.companies).toContainEqual(expect.objectContaining({ stock_code: '2330', name: '台積電' }));
      expect(exported.data.material_events).toContainEqual(expect.objectContaining({ id: event.id, deleted_at: now }));
      expect(exported.data.material_event_deletion_audit).toContainEqual(expect.objectContaining({ event_id: event.id, deleted_at: now }));
      expect(exported.settings).toContainEqual(expect.objectContaining({ key: 'monitoring', value: { enabled: true } }));
      expect(exported.settings).toContainEqual(expect.objectContaining({ key: 'safe-list', value: ['新聞', 3, null] }));
      expect(serialized).not.toContain('ACCESS-TOKEN-MUST-NOT-EXPORT');
      expect(serialized).not.toContain('REFRESH-TOKEN-MUST-NOT-EXPORT');
      expect(serialized).not.toContain('CLIENT-SECRET-MUST-NOT-EXPORT');
      expect(serialized).not.toContain('ARRAY-TOKEN-MUST-NOT-EXPORT');
      expect(exported.settings.some(({ key }) => key === 'test-only-sensitive-value')).toBe(false);
      expect(exported.settings.some(({ key }) => key === 'nested-secret-list')).toBe(false);
    } finally {
      database.close();
    }
  });
});
