import type { classifyConferenceSchedule } from './conference-task-plan';
import type { createLegacyConfigBridge } from './legacy-config-bridge';

type SchedulePlan = ReturnType<typeof classifyConferenceSchedule>;
type Bridge = ReturnType<typeof createLegacyConfigBridge>;

export function createConferenceTaskController(dependencies: {
  inspectPlan(): Promise<SchedulePlan>;
  ensure(setupReady: boolean): Promise<Exclude<SchedulePlan, { kind: 'create-v2' }>>;
  switchFromDisabledLegacy?(setupReady: boolean): Promise<Exclude<SchedulePlan, { kind: 'create-v2' }>>;
  isSetupReady(): Promise<boolean>;
  bridge: Pick<Bridge, 'review' | 'confirm'>;
}) {
  let pendingEnsure: Promise<SchedulePlan & { setupReady: boolean }> | null = null;
  let pendingSwitch: Promise<SchedulePlan & { setupReady: boolean }> | null = null;
  return {
    async status() {
      const [plan, setupReady] = await Promise.all([dependencies.inspectPlan(), dependencies.isSetupReady()]);
      return { ...plan, setupReady };
    },
    ensure() {
      if (pendingEnsure) return pendingEnsure;
      pendingEnsure = (async () => {
        const setupReady = await dependencies.isSetupReady();
        return { ...await dependencies.ensure(setupReady), setupReady };
      })().finally(() => { pendingEnsure = null; });
      return pendingEnsure;
    },
    switchFromDisabledLegacy() {
      if (pendingSwitch) return pendingSwitch;
      pendingSwitch = (async () => {
        const setupReady = await dependencies.isSetupReady();
        if (!dependencies.switchFromDisabledLegacy) return { kind: 'blocked' as const, reason: '此版本尚未支援切換每日工作', setupReady };
        return { ...await dependencies.switchFromDisabledLegacy(setupReady), setupReady };
      })().finally(() => { pendingSwitch = null; });
      return pendingSwitch;
    },
    reviewLegacy: () => dependencies.bridge.review(),
    confirmLegacy: (hash: string) => dependencies.bridge.confirm(hash),
  };
}
