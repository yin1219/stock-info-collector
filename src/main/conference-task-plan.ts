import path from 'node:path';

export const LEGACY_CONFERENCE_TASK_NAME = 'stock-info-collector-daily-work';
export const ALTERNATE_LEGACY_CONFERENCE_TASK_NAME = 'stock-info-collector';
export const V2_CONFERENCE_TASK_NAME = 'StockReporterAssistant-ConferenceDaily';
export const V2_CONFERENCE_TASK_ARGUMENTS = '--processStart StockReporterAssistant.exe --process-start-args "--conference-scheduled"';

export interface ConferenceTaskInfo {
  name: string;
  execute: string;
  workingDirectory: string;
  arguments: string;
  enabled: boolean;
  running?: boolean;
  taskState?: number;
  dailyAtLocal?: string | null;
  multipleInstances?: string;
  lastTaskResult?: number;
}

type Inspection = ConferenceTaskInfo | null | 'inaccessible';
type SchedulePlan = { kind: 'legacy' | 'legacy-disabled'; configPath: string; lastTaskResult?: number }
  | { kind: 'v2'; legacyConfigPath?: string; lastTaskResult?: number } | { kind: 'create-v2' }
  | { kind: 'upgrade-v2-time' }
  | { kind: 'blocked'; reason: string };

function samePath(left: string, right: string): boolean {
  return path.win32.resolve(left).toLowerCase() === path.win32.resolve(right).toLowerCase();
}

export function classifyConferenceSchedule(input: {
  legacy: Inspection;
  alternateLegacy?: Inspection;
  current: Inspection;
  expectedUpdateExe: string;
  legacyConfigExists: boolean;
}): SchedulePlan {
  if (input.legacy === 'inaccessible' || input.alternateLegacy === 'inaccessible' || input.current === 'inaccessible') {
    return { kind: 'blocked', reason: '無法確認現有 Windows 法說會工作，未建立第二個排程' };
  }
  for (const [task, name] of [[input.legacy, LEGACY_CONFERENCE_TASK_NAME],
    [input.alternateLegacy, ALTERNATE_LEGACY_CONFERENCE_TASK_NAME]] as const) {
    if (!task) continue;
    if (task.name !== name || !task.dailyAtLocal
      || task.arguments.trim() !== ''
      || !path.win32.isAbsolute(task.execute)
      || path.win32.basename(task.execute).toLowerCase() !== 'stock-info-collector.exe'
      || !samePath(path.win32.dirname(task.execute), task.workingDirectory)
      || !input.legacyConfigExists) {
      return { kind: 'blocked', reason: '既有法說會工作或設定無法辨識，未覆蓋原工作' };
    }
  }
  if (input.legacy && input.alternateLegacy
    && !samePath(input.legacy.workingDirectory, input.alternateLegacy.workingDirectory)) {
    return { kind: 'blocked', reason: '多個 v1 法說會工作的設定位置不同，未確認唯一寫入者' };
  }
  const legacy = input.legacy ?? input.alternateLegacy ?? null;
  const otherLegacy = input.legacy && input.alternateLegacy ? input.alternateLegacy : null;
  if (otherLegacy && (otherLegacy.enabled || otherLegacy.running !== false || otherLegacy.taskState !== 1)) {
    return { kind: 'blocked', reason: '同時存在可能執行的 v1 與 v2 法說會工作，需先確認唯一寫入者' };
  }
  if (input.current) {
    const task = input.current;
    if (task.name !== V2_CONFERENCE_TASK_NAME
      || !samePath(task.execute, input.expectedUpdateExe)
      || !samePath(task.workingDirectory, path.win32.dirname(input.expectedUpdateExe))
      || task.arguments !== V2_CONFERENCE_TASK_ARGUMENTS
      || task.multipleInstances !== 'IgnoreNew') {
      return { kind: 'blocked', reason: '同名 Windows 工作不是 v2 建立的法說會工作，未修改' };
    }
    if (task.dailyAtLocal === '17:45' && task.enabled
      && (!legacy || (!legacy.enabled && legacy.running === false && legacy.taskState === 1))) {
      return { kind: 'upgrade-v2-time' };
    }
    if (task.dailyAtLocal !== '19:00') {
      const actualTime = /^\d{2}:\d{2}$/.test(task.dailyAtLocal ?? '') ? task.dailyAtLocal : '未知';
      return { kind: 'blocked', reason: `v2 法說會工作的每日時間為 ${actualTime}，預期 19:00；目前不會執行每日同步，也未自動修改 Windows 工作` };
    }
    if (!task.enabled) {
      return { kind: 'blocked', reason: '偵測到已停用的 v2 法說會工作；目前沒有啟用中的每日法說會排程，未自動修改 Windows 工作' };
    }
    if (legacy && (legacy.enabled || legacy.running !== false || legacy.taskState !== 1)) {
      return { kind: 'blocked', reason: '同時存在可能執行的 v1 與 v2 法說會工作，需先確認唯一寫入者' };
    }
    return {
      kind: 'v2',
      ...(legacy ? { legacyConfigPath: path.win32.join(legacy.workingDirectory, 'config', 'default.json') } : {}),
      ...(task.lastTaskResult === undefined ? {} : { lastTaskResult: task.lastTaskResult }),
    };
  }
  if (legacy) {
    if (!legacy.enabled && (legacy.running !== false || legacy.taskState !== 1)) {
      return { kind: 'blocked', reason: '無法確認已停用的 v1 法說會工作是否仍在執行，未啟動 Calendar 同步' };
    }
    return {
      kind: legacy.enabled ? 'legacy' : 'legacy-disabled',
      configPath: path.win32.join(legacy.workingDirectory, 'config', 'default.json'),
      ...(legacy.lastTaskResult === undefined ? {} : { lastTaskResult: legacy.lastTaskResult }),
    };
  }
  return { kind: 'create-v2' };
}
