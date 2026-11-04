import type { WatchlistIpcService } from './watchlist-ipc';
import type { classifyConferenceSchedule } from './conference-task-plan';

type SchedulePlan = ReturnType<typeof classifyConferenceSchedule>;

export function createCoexistingWatchlistController(dependencies: {
  base: WatchlistIpcService;
  inspectPlan(): Promise<SchedulePlan>;
  applyIfConfirmed(): Promise<{ status: string }>;
  afterMutation?(): Promise<void>;
}): WatchlistIpcService {
  const { base } = dependencies;

  async function propagate(): Promise<void> {
    try {
      const plan = await dependencies.inspectPlan();
      if (plan.kind === 'create-v2' || (plan.kind === 'v2' && !plan.legacyConfigPath)) return;
      if (plan.kind === 'blocked') throw new Error(plan.reason);
      const result = await dependencies.applyIfConfirmed();
      if (result.status === 'not-confirmed') {
        throw new Error('v1 股號尚未核對確認，舊版工作未同步');
      }
    } catch (cause) {
      throw new Error(`v2 清單已更新，但 v1 法說會工作未同步：${cause instanceof Error ? cause.message : String(cause)}`);
    }
  }

  return {
    list: (filter) => base.list(filter),
    async add(input) {
      const result = await base.add(input);
      await propagate();
      await dependencies.afterMutation?.();
      return result;
    },
    update: (input) => base.update(input),
    async setActive(input) {
      const result = await base.setActive(input);
      await propagate();
      await dependencies.afterMutation?.();
      return result;
    },
    async remove(input) {
      const result = await base.remove(input);
      await propagate();
      await dependencies.afterMutation?.();
      return result;
    },
    previewImport: (input) => base.previewImport(input),
    async applyImport(input) {
      const result = await base.applyImport(input);
      await propagate();
      await dependencies.afterMutation?.();
      return result;
    },
  };
}
