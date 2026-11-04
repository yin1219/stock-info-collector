import path from 'node:path';
import type { createRepositories } from '../repositories';
import { inspectLegacyConfig, syncLegacyStockNumbers } from '../services/legacy-config-sync';
import type { classifyConferenceSchedule } from './conference-task-plan';

type Repositories = ReturnType<typeof createRepositories>;
type Confirmation = { configPath: string; hash: string; confirmedAt: string };
const settingKey = 'legacy-v1-watchlist-sync';

export function legacyConfigPathForPlan(plan: ReturnType<typeof classifyConferenceSchedule>): string | null {
  if (plan.kind === 'legacy' || plan.kind === 'legacy-disabled') return plan.configPath;
  return plan.kind === 'v2' ? plan.legacyConfigPath ?? null : null;
}

function samePath(left: string, right: string): boolean {
  return path.win32.normalize(left).toLowerCase() === path.win32.normalize(right).toLowerCase();
}

export function createLegacyConfigBridge(dependencies: {
  repositories: Repositories;
  locateLegacyConfig(): Promise<string | null>;
  now?: () => string;
}) {
  const now = dependencies.now ?? (() => new Date().toISOString());
  const activeCodes = () => [...new Set(dependencies.repositories.watchlist.list({ activeOnly: true })
    .map(({ stockCode }) => stockCode))];
  async function recognizedPath(): Promise<string> {
    const configPath = await dependencies.locateLegacyConfig();
    if (!configPath) throw new Error('無法辨識既有 v1 法說會工作或設定，未套用關注股號');
    return configPath;
  }
  return {
    async review() {
      const configPath = await recognizedPath();
      const config = await inspectLegacyConfig(configPath);
      const confirmation = dependencies.repositories.settings.get<Confirmation>(settingKey);
      return {
        configPath, hash: config.hash, currentCodes: config.stockNumbers,
        activeCodes: activeCodes(),
        confirmed: Boolean(confirmation && samePath(confirmation.configPath, configPath)
          && confirmation.hash === config.hash),
      };
    },
    async confirm(reviewedHash: string) {
      const configPath = await recognizedPath();
      const config = await inspectLegacyConfig(configPath);
      if (config.hash !== reviewedHash) throw new Error('v1 設定已變更，請重新核對後再套用');
      const result = await syncLegacyStockNumbers({
        configPath, reviewedHash, activeStockCodes: activeCodes(),
      });
      const updated = await inspectLegacyConfig(configPath);
      dependencies.repositories.settings.set(settingKey, { configPath, hash: updated.hash, confirmedAt: now() }, now());
      return result;
    },
    async applyIfConfirmed() {
      const confirmation = dependencies.repositories.settings.get<Confirmation>(settingKey);
      if (!confirmation) return { status: 'not-confirmed' as const };
      const configPath = await recognizedPath();
      if (!samePath(configPath, confirmation.configPath)) {
        throw new Error('已確認的 v1 工作設定路徑改變，請重新核對後再套用');
      }
      const config = await inspectLegacyConfig(configPath);
      if (config.hash !== confirmation.hash) throw new Error('v1 設定已被外部修改，請重新核對後再套用');
      const result = await syncLegacyStockNumbers({
        configPath, reviewedHash: confirmation.hash, activeStockCodes: activeCodes(),
      });
      const updated = await inspectLegacyConfig(configPath);
      dependencies.repositories.settings.set(settingKey, { ...confirmation, hash: updated.hash }, now());
      return result;
    },
  };
}
