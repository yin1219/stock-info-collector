import { describe, expect, it } from 'vitest';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createCoexistingWatchlistController } from '../../src/main/coexisting-watchlist-controller';
import { createConferenceTaskController } from '../../src/main/conference-task-controller';
import { registerConferenceTaskIpc } from '../../src/main/conference-task-ipc';
import { classifyConferenceSchedule, LEGACY_CONFERENCE_TASK_NAME, V2_CONFERENCE_TASK_ARGUMENTS, V2_CONFERENCE_TASK_NAME } from '../../src/main/conference-task-plan';
import { createLegacyConfigBridge, legacyConfigPathForPlan } from '../../src/main/legacy-config-bridge';
import { createWatchlistController } from '../../src/main/watchlist-controller';
import { registerWatchlistIpc } from '../../src/main/watchlist-ipc';
import { createRepositories } from '../../src/repositories';
import { openDatabase } from '../../src/repositories/migrations';
import { createTestPaths } from '../support';

describe('watchlist-management / retained v1 rollback config / main IPC and SQLite integration', () => {
  it('reviews, confirms, and follows a v2 watchlist change while preserving the disabled v1 config backup', async () => {
    const paths = await createTestPaths();
    const database = openDatabase(paths.database);
    try {
      const repositories = createRepositories(database);
      const legacyRoot = path.join(paths.root, 'legacy');
      const configPath = path.join(legacyRoot, 'config', 'default.json');
      await mkdir(path.dirname(configPath), { recursive: true });
      await writeFile(configPath, '{"Debug":false,"LogLevel":"info","StockNumbers":"2330","Extra":{"keep":true}}\n', 'utf8');
      const updateExe = path.join(paths.root, 'installed-v2', 'Update.exe');
      const inspectPlan = async () => classifyConferenceSchedule({
        legacy: {
          name: LEGACY_CONFERENCE_TASK_NAME, execute: path.join(legacyRoot, 'stock-info-collector.exe'),
          workingDirectory: legacyRoot, arguments: '', enabled: false, running: false,
          taskState: 1, dailyAtLocal: '19:00',
        },
        current: {
          name: V2_CONFERENCE_TASK_NAME, execute: updateExe, workingDirectory: path.dirname(updateExe),
          arguments: V2_CONFERENCE_TASK_ARGUMENTS, enabled: true, dailyAtLocal: '19:00',
          multipleInstances: 'IgnoreNew',
        },
        expectedUpdateExe: updateExe, legacyConfigExists: true,
      });
      const bridge = createLegacyConfigBridge({
        repositories, locateLegacyConfig: async () => legacyConfigPathForPlan(await inspectPlan()),
        now: () => '2026-11-02T09:00:00.000Z',
      });
      const controller = createConferenceTaskController({
        inspectPlan,
        isSetupReady: async () => true,
        ensure: async () => ({ kind: 'blocked', reason: 'not used' }),
        bridge,
      });
      const base = createWatchlistController({
        repositories,
        directory: { async findByStockCode(market, stockCode) {
          return { market, stockCode, name: `測試公司 ${stockCode}` };
        } },
        now: () => '2026-11-02T09:00:00.000Z',
      });
      const initial = repositories.companies.upsert({
        market: 'TWSE', stockCode: '2454', name: '測試公司 2454', updatedAt: '2026-11-02T09:00:00.000Z',
      });
      repositories.watchlist.upsert({
        companyId: initial.id, active: true, category: '', notes: '',
        createdAt: '2026-11-02T09:00:00.000Z', updatedAt: '2026-11-02T09:00:00.000Z',
      });

      const handlers = new Map<string, (event: unknown, payload?: unknown) => unknown>();
      const ipc = {
        handle(channel: string, handler: (event: unknown, payload?: unknown) => unknown) { handlers.set(channel, handler); },
        removeHandler(channel: string) { handlers.delete(channel); },
      };
      const stopTaskIpc = registerConferenceTaskIpc(ipc, controller, () => true);
      const stopWatchlistIpc = registerWatchlistIpc(ipc, createCoexistingWatchlistController({
        base, inspectPlan, applyIfConfirmed: bridge.applyIfConfirmed,
      }), () => true);
      const invoke = async (channel: string, payload?: unknown) => {
        const handler = handlers.get(channel);
        if (!handler) throw new Error(`Missing IPC handler: ${channel}`);
        return handler({}, payload);
      };
      try {
        expect(await invoke('conference-task:status')).toMatchObject({ kind: 'v2', legacyConfigPath: configPath });
        const review = await invoke('conference-task:review-legacy') as { hash: string; currentCodes: string[]; activeCodes: string[] };
        expect(review).toMatchObject({ currentCodes: ['2330'], activeCodes: ['2454'] });
        expect(JSON.parse(await readFile(configPath, 'utf8')).StockNumbers).toBe('2330');
        const confirmed = await invoke('conference-task:confirm-legacy', { hash: review.hash }) as { backupPath: string };
        expect(await readFile(confirmed.backupPath, 'utf8')).toContain('"StockNumbers":"2330"');
        expect(JSON.parse(await readFile(configPath, 'utf8'))).toMatchObject({
          Debug: false, LogLevel: 'info', StockNumbers: '2454', Extra: { keep: true },
        });
        await invoke('watchlist:add', { market: 'TPEX', stockCode: '6488' });
        expect(JSON.parse(await readFile(configPath, 'utf8')).StockNumbers).toBe('2454,6488');
      } finally {
        stopTaskIpc();
        stopWatchlistIpc();
      }
    } finally {
      database.close();
      await paths.cleanup();
    }
  });
});
