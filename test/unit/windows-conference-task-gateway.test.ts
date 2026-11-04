import { describe, expect, it, vi } from 'vitest';
import { createWindowsConferenceTaskGateway } from '../../src/main/windows-conference-task-gateway';
import { LEGACY_CONFERENCE_TASK_NAME, V2_CONFERENCE_TASK_NAME } from '../../src/main/conference-task-plan';

const scriptPath = 'C:\\App\\resources\\conference-task.ps1';

describe('windows-distribution / conference task gateway / fail-closed Windows boundary', () => {
  it('parses a recognized task and never constructs a shell command from paths', async () => {
    const execute = vi.fn(async (_file: string, _args: string[]) => JSON.stringify({
      status: 'found', task: {
        name: LEGACY_CONFERENCE_TASK_NAME, execute: 'F:\\Old Reporter\\stock-info-collector.exe',
        workingDirectory: 'F:\\Old Reporter', arguments: '', enabled: true, running: false, taskState: 3,
        dailyAtLocal: '19:00', multipleInstances: 'IgnoreNew', lastTaskResult: 0x800710e0,
      },
    }));
    const gateway = createWindowsConferenceTaskGateway({ scriptPath, execute });
    expect(await gateway.inspect(LEGACY_CONFERENCE_TASK_NAME)).toMatchObject({
      execute: 'F:\\Old Reporter\\stock-info-collector.exe', running: false, taskState: 3, lastTaskResult: 0x800710e0,
    });
    expect(execute).toHaveBeenCalledWith('powershell.exe', [
      '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
      '-Action', 'Inspect', '-TaskName', LEGACY_CONFERENCE_TASK_NAME,
    ]);
  });

  it('treats only an explicit missing response as absent', async () => {
    const missing = createWindowsConferenceTaskGateway({ scriptPath, execute: async () => '{"status":"missing"}' });
    expect(await missing.inspect(V2_CONFERENCE_TASK_NAME)).toBeNull();
    const denied = createWindowsConferenceTaskGateway({ scriptPath, execute: async () => { throw new Error('access denied'); } });
    expect(await denied.inspect(V2_CONFERENCE_TASK_NAME)).toBe('inaccessible');
    const malformed = createWindowsConferenceTaskGateway({ scriptPath, execute: async () => '{"status":"found","task":{}}' });
    expect(await malformed.inspect(V2_CONFERENCE_TASK_NAME)).toBe('inaccessible');
    const missingRunningState = createWindowsConferenceTaskGateway({ scriptPath, execute: async () => JSON.stringify({
      status: 'found', task: {
        name: LEGACY_CONFERENCE_TASK_NAME, execute: 'F:\\Old Reporter\\stock-info-collector.exe',
        workingDirectory: 'F:\\Old Reporter', arguments: '', enabled: false, dailyAtLocal: '19:00',
      },
    }) });
    expect(await missingRunningState.inspect(LEGACY_CONFERENCE_TASK_NAME)).toBe('inaccessible');
  });

  it('allows read-only inspection of the older v1 task name', async () => {
    const execute = vi.fn(async () => '{"status":"missing"}');
    const gateway = createWindowsConferenceTaskGateway({ scriptPath, execute });
    expect(await gateway.inspect('stock-info-collector')).toBeNull();
    expect(execute).toHaveBeenCalledWith('powershell.exe', expect.arrayContaining([
      '-Action', 'Inspect', '-TaskName', 'stock-info-collector',
    ]));
  });

  it('requires an explicit registered response and passes arguments without shell interpolation', async () => {
    const execute = vi.fn(async () => '{"status":"registered"}');
    const gateway = createWindowsConferenceTaskGateway({ scriptPath, execute });
    const registration = {
      name: V2_CONFERENCE_TASK_NAME, execute: 'C:\\App\\Update.exe', workingDirectory: 'C:\\App',
      arguments: '--processStart StockReporterAssistant.exe --process-start-args "--conference-scheduled"',
      localTime: '19:00' as const, multipleInstances: 'IgnoreNew' as const,
    };
    await gateway.registerIfAbsent(registration);
    expect(execute).toHaveBeenCalledWith('powershell.exe', [
      '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
      '-Action', 'Register', '-TaskName', V2_CONFERENCE_TASK_NAME,
      '-Executable', registration.execute, '-WorkingDirectory', registration.workingDirectory,
      '-Arguments', registration.arguments,
    ]);
    const unconfirmed = createWindowsConferenceTaskGateway({ scriptPath, execute: async () => '{"status":"missing"}' });
    await expect(unconfirmed.registerIfAbsent(registration)).rejects.toThrow(/無法確認/);
  });

  it('requests deletion only of the exact v2-owned task and accepts an explicit skipped result', async () => {
    const execute = vi.fn(async () => '{"status":"removed"}');
    const gateway = createWindowsConferenceTaskGateway({ scriptPath, execute });
    await expect(gateway.removeIfOwned('C:\\App\\Update.exe')).resolves.toBe('removed');
    expect(execute).toHaveBeenCalledWith('powershell.exe', [
      '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
      '-Action', 'Remove', '-TaskName', V2_CONFERENCE_TASK_NAME,
      '-Executable', 'C:\\App\\Update.exe',
    ]);
    const skipped = createWindowsConferenceTaskGateway({ scriptPath, execute: async () => '{"status":"skipped"}' });
    await expect(skipped.removeIfOwned('C:\\App\\Update.exe')).resolves.toBe('skipped');
    await expect(gateway.removeIfOwned('C:\\App\\Other.exe')).rejects.toThrow(/Update\.exe/);
  });

  it('passes the disabled-legacy exception only for an explicitly confirmed switch', async () => {
    const execute = vi.fn(async () => '{"status":"registered"}');
    const gateway = createWindowsConferenceTaskGateway({ scriptPath, execute });
    await gateway.registerIfAbsent({
      name: V2_CONFERENCE_TASK_NAME, execute: 'C:\\App\\Update.exe', workingDirectory: 'C:\\App',
      arguments: '--processStart StockReporterAssistant.exe --process-start-args "--conference-scheduled"',
      localTime: '19:00', multipleInstances: 'IgnoreNew', allowDisabledLegacy: true,
    });
    expect(execute.mock.calls[0][1]).toContain('-AllowDisabledLegacy');
  });

  it('requests a guarded 17:45 to 19:00 migration with a userData backup directory', async () => {
    const execute = vi.fn(async () => '{"status":"migrated"}');
    const gateway = createWindowsConferenceTaskGateway({
      scriptPath, execute, backupDirectory: 'C:\\Users\\Test\\AppData\\Roaming\\StockReporterAssistant\\windows-task-backups',
    });
    await expect(gateway.migrateOwnedTime('C:\\App\\Update.exe')).resolves.toBeUndefined();
    expect(execute).toHaveBeenCalledWith('powershell.exe', [
      '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath,
      '-Action', 'MigrateTime', '-TaskName', V2_CONFERENCE_TASK_NAME,
      '-Executable', 'C:\\App\\Update.exe',
      '-BackupDirectory', 'C:\\Users\\Test\\AppData\\Roaming\\StockReporterAssistant\\windows-task-backups',
    ]);
    await expect(gateway.migrateOwnedTime('C:\\App\\other.exe')).rejects.toThrow(/Update\.exe/);
  });
});
