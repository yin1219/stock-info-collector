import path from 'node:path';
import {
  classifyConferenceSchedule,
  ALTERNATE_LEGACY_CONFERENCE_TASK_NAME,
  LEGACY_CONFERENCE_TASK_NAME,
  V2_CONFERENCE_TASK_ARGUMENTS,
  V2_CONFERENCE_TASK_NAME,
  type ConferenceTaskInfo,
} from './conference-task-plan';

type Inspection = ConferenceTaskInfo | null | 'inaccessible';

export interface ConferenceTaskRegistration {
  name: string;
  execute: string;
  workingDirectory: string;
  arguments: string;
  localTime: '19:00';
  multipleInstances: 'IgnoreNew';
  allowDisabledLegacy?: boolean;
}

export async function bootstrapConferenceTask(dependencies: {
  inspect(name: string): Promise<Inspection>;
  legacyConfigExists(configPath: string): Promise<boolean>;
  registerIfAbsent(registration: ConferenceTaskRegistration): Promise<void>;
  migrateOwnedTime?(expectedUpdateExe: string): Promise<void>;
  expectedUpdateExe: string;
  setupReady: boolean;
  switchDisabledLegacy?: boolean;
}): Promise<{ kind: 'legacy' | 'legacy-disabled'; configPath: string } | { kind: 'v2'; legacyConfigPath?: string } | { kind: 'blocked'; reason: string }> {
  if (!path.win32.isAbsolute(dependencies.expectedUpdateExe)
    || path.win32.basename(dependencies.expectedUpdateExe).toLowerCase() !== 'update.exe') {
    return { kind: 'blocked', reason: '找不到跨版本穩定的 Windows 更新程式，未建立法說會排程' };
  }

  async function inspectPlan() {
    const [legacy, alternateLegacy, current] = await Promise.all([
      dependencies.inspect(LEGACY_CONFERENCE_TASK_NAME),
      dependencies.inspect(ALTERNATE_LEGACY_CONFERENCE_TASK_NAME),
      dependencies.inspect(V2_CONFERENCE_TASK_NAME),
    ]);
    const knownLegacy = legacy && legacy !== 'inaccessible' ? legacy : alternateLegacy;
    const legacyConfigPath = knownLegacy && knownLegacy !== 'inaccessible'
      ? path.win32.join(knownLegacy.workingDirectory, 'config', 'default.json') : '';
    const legacyConfigExists = legacyConfigPath
      ? await dependencies.legacyConfigExists(legacyConfigPath) : false;
    return classifyConferenceSchedule({
      legacy, alternateLegacy, current, expectedUpdateExe: dependencies.expectedUpdateExe, legacyConfigExists,
    });
  }

  try {
    const plan = await inspectPlan();
    if (plan.kind === 'blocked' || plan.kind === 'legacy' || plan.kind === 'v2') return plan;
    if (plan.kind === 'upgrade-v2-time') {
      if (!dependencies.migrateOwnedTime) {
        return { kind: 'blocked', reason: 'v2 每日工作仍為 17:45，無法安全遷移至 19:00' };
      }
      await dependencies.migrateOwnedTime(dependencies.expectedUpdateExe);
      const verified = await inspectPlan();
      return verified.kind === 'v2' ? verified
        : { kind: 'blocked', reason: 'v2 每日工作遷移後無法確認 19:00 與唯一寫入者' };
    }
    if (dependencies.switchDisabledLegacy && plan.kind !== 'legacy-disabled') {
      return { kind: 'blocked', reason: '沒有已停用且停止的 v1 工作可切換，未建立 v2 工作' };
    }
    if (!dependencies.switchDisabledLegacy && plan.kind !== 'create-v2') return plan;
    if (!dependencies.setupReady) {
      return { kind: 'blocked', reason: '請先完成 Google 連線與關注公司設定，再建立法說會每日工作' };
    }
    await dependencies.registerIfAbsent({
      name: V2_CONFERENCE_TASK_NAME,
      execute: dependencies.expectedUpdateExe,
      workingDirectory: path.win32.dirname(dependencies.expectedUpdateExe),
      arguments: V2_CONFERENCE_TASK_ARGUMENTS,
      localTime: '19:00',
      multipleInstances: 'IgnoreNew',
      ...(dependencies.switchDisabledLegacy ? { allowDisabledLegacy: true } : {}),
    });
    const verified = await inspectPlan();
    return verified.kind === 'v2'
      ? verified
      : { kind: 'blocked', reason: 'Windows 法說會工作建立後無法確認唯一寫入者，請檢查工作排程器' };
  } catch {
    return { kind: 'blocked', reason: 'Windows 法說會工作無法確認或建立，未變更既有工作' };
  }
}
