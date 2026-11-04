import { execFile } from 'node:child_process';
import path from 'node:path';
import {
  ALTERNATE_LEGACY_CONFERENCE_TASK_NAME,
  LEGACY_CONFERENCE_TASK_NAME,
  V2_CONFERENCE_TASK_ARGUMENTS,
  V2_CONFERENCE_TASK_NAME,
  type ConferenceTaskInfo,
} from './conference-task-plan';
import type { ConferenceTaskRegistration } from './conference-task-bootstrap';

type Execute = (file: string, arguments_: string[]) => Promise<string>;

function runPowerShell(file: string, arguments_: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile(file, arguments_, { windowsHide: true, timeout: 15_000, maxBuffer: 64 * 1024 }, (error, stdout) => {
      if (error) reject(error);
      else resolve(stdout);
    });
  });
}

function parseResponse(output: string): Record<string, unknown> {
  const response: unknown = JSON.parse(output.trim());
  if (!response || typeof response !== 'object' || Array.isArray(response)) {
    throw new Error('Windows 工作排程器回應格式無效');
  }
  return response as Record<string, unknown>;
}

function taskInfo(value: unknown, requestedName: string): ConferenceTaskInfo {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Windows 工作資訊無效');
  const task = value as Record<string, unknown>;
  if (task.name !== requestedName || typeof task.execute !== 'string'
    || typeof task.workingDirectory !== 'string' || typeof task.arguments !== 'string'
    || typeof task.enabled !== 'boolean' || typeof task.running !== 'boolean'
    || !Number.isInteger(task.taskState) || (task.taskState as number) < 0 || (task.taskState as number) > 4
    || task.lastTaskResult !== undefined && (!Number.isInteger(task.lastTaskResult)
      || (task.lastTaskResult as number) < 0 || (task.lastTaskResult as number) > 0xffffffff)) {
    throw new Error('Windows 工作資訊不完整');
  }
  return task as unknown as ConferenceTaskInfo;
}

export function createWindowsConferenceTaskGateway(dependencies: {
  scriptPath: string;
  execute?: Execute;
  backupDirectory?: string;
}) {
  if (!path.win32.isAbsolute(dependencies.scriptPath)) throw new Error('Windows 工作排程器腳本路徑必須是絕對路徑');
  const execute = dependencies.execute ?? runPowerShell;
  const baseArguments = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', dependencies.scriptPath];
  return {
    async inspect(name: string): Promise<ConferenceTaskInfo | null | 'inaccessible'> {
      if (name !== LEGACY_CONFERENCE_TASK_NAME && name !== ALTERNATE_LEGACY_CONFERENCE_TASK_NAME
        && name !== V2_CONFERENCE_TASK_NAME) return 'inaccessible';
      try {
        const response = parseResponse(await execute('powershell.exe', [...baseArguments,
          '-Action', 'Inspect', '-TaskName', name,
        ]));
        if (response.status === 'missing') return null;
        if (response.status === 'found') return taskInfo(response.task, name);
      } catch {
        // Missing is accepted only when the Task Scheduler script explicitly reports it.
      }
      return 'inaccessible';
    },
    async registerIfAbsent(registration: ConferenceTaskRegistration): Promise<void> {
      if (registration.name !== V2_CONFERENCE_TASK_NAME
        || registration.arguments !== V2_CONFERENCE_TASK_ARGUMENTS
        || registration.localTime !== '19:00'
        || registration.multipleInstances !== 'IgnoreNew') {
        throw new Error('拒絕建立非 v2 法說會每日工作');
      }
      const response = parseResponse(await execute('powershell.exe', [...baseArguments,
        '-Action', 'Register', '-TaskName', registration.name,
        '-Executable', registration.execute, '-WorkingDirectory', registration.workingDirectory,
        '-Arguments', registration.arguments,
        ...(registration.allowDisabledLegacy ? ['-AllowDisabledLegacy'] : []),
      ]));
      if (response.status !== 'registered') throw new Error('Windows 法說會工作建立後無法確認');
    },
    async migrateOwnedTime(expectedUpdateExe: string): Promise<void> {
      if (!path.win32.isAbsolute(expectedUpdateExe)
        || path.win32.basename(expectedUpdateExe).toLowerCase() !== 'update.exe'
        || !dependencies.backupDirectory || !path.win32.isAbsolute(dependencies.backupDirectory)) {
        throw new Error('只能遷移由 v2 Update.exe 管理的工作，且須提供絕對備份路徑');
      }
      const response = parseResponse(await execute('powershell.exe', [...baseArguments,
        '-Action', 'MigrateTime', '-TaskName', V2_CONFERENCE_TASK_NAME,
        '-Executable', expectedUpdateExe, '-BackupDirectory', dependencies.backupDirectory,
      ]));
      if (response.status !== 'migrated') throw new Error('v2 每日工作遷移結果無法確認');
    },
    async removeIfOwned(expectedUpdateExe: string): Promise<'removed' | 'skipped'> {
      if (!path.win32.isAbsolute(expectedUpdateExe)
        || path.win32.basename(expectedUpdateExe).toLowerCase() !== 'update.exe') {
        throw new Error('只能移除由 v2 Update.exe 管理的法說會工作');
      }
      const response = parseResponse(await execute('powershell.exe', [...baseArguments,
        '-Action', 'Remove', '-TaskName', V2_CONFERENCE_TASK_NAME,
        '-Executable', expectedUpdateExe,
      ]));
      if (response.status === 'removed' || response.status === 'skipped') return response.status;
      throw new Error('Windows 法說會工作移除狀態無法確認');
    },
  };
}
