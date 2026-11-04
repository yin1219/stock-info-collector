import { describe, expect, it } from 'vitest';
import { classifyConferenceSchedule } from '../../src/main/conference-task-plan';

const v1 = {
  name: 'stock-info-collector-daily-work',
  execute: 'F:\\OldReporter\\stock-info-collector.exe',
  workingDirectory: 'F:\\OldReporter',
  arguments: '', enabled: true, dailyAtLocal: '19:00',
};
const v2 = {
  name: 'StockReporterAssistant-ConferenceDaily',
  execute: 'C:\\Users\\Test\\AppData\\Local\\StockReporterAssistant\\Update.exe',
  workingDirectory: 'C:\\Users\\Test\\AppData\\Local\\StockReporterAssistant',
  arguments: '--processStart StockReporterAssistant.exe --process-start-args "--conference-scheduled"',
  enabled: true, dailyAtLocal: '19:00', multipleInstances: 'IgnoreNew',
};
const expectedUpdateExe = v2.execute;

describe('windows-distribution / conference task bootstrap / one daily Calendar writer', () => {
  it('preserves a recognized v1 task and its config instead of creating a second writer', () => {
    expect(classifyConferenceSchedule({ legacy: v1, current: null, expectedUpdateExe, legacyConfigExists: true }))
      .toEqual({ kind: 'legacy', configPath: 'F:\\OldReporter\\config\\default.json' });
  });

  it('keeps a failed v1 task as the sole owner while exposing its last Windows result', () => {
    expect(classifyConferenceSchedule({
      legacy: { ...v1, lastTaskResult: 0x800710e0 }, current: null, expectedUpdateExe, legacyConfigExists: true,
    })).toEqual({ kind: 'legacy', configPath: 'F:\\OldReporter\\config\\default.json', lastTaskResult: 0x800710e0 });
  });

  it('permits only manual acceptance while a recognized v1 task is disabled and not running', () => {
    expect(classifyConferenceSchedule({
      legacy: { ...v1, enabled: false, running: false, taskState: 1 }, current: null, expectedUpdateExe, legacyConfigExists: true,
    })).toEqual({
      kind: 'legacy-disabled', configPath: 'F:\\OldReporter\\config\\default.json',
    });
  });

  it('does not permit manual acceptance when a disabled v1 task is running or its state is unknown', () => {
    for (const [running, taskState] of [[true, 4], [undefined, 1], [false, 0]] as const) {
      expect(classifyConferenceSchedule({
        legacy: { ...v1, enabled: false, running, taskState }, current: null, expectedUpdateExe, legacyConfigExists: true,
      }).kind).toBe('blocked');
    }
  });

  it('accepts v2 as the sole active writer while a recognized disabled v1 task is retained for rollback', () => {
    expect(classifyConferenceSchedule({
      legacy: { ...v1, enabled: false, running: false, taskState: 1 }, current: v2,
      expectedUpdateExe, legacyConfigExists: true,
    })).toEqual({ kind: 'v2', legacyConfigPath: 'F:\\OldReporter\\config\\default.json' });
    expect(classifyConferenceSchedule({
      legacy: { ...v1, enabled: true, running: false }, current: v2,
      expectedUpdateExe, legacyConfigExists: true,
    }).kind).toBe('blocked');
  });

  it('blocks v2 when the older 18:15 v1 task is still enabled even if the known v1 task is disabled', () => {
    expect(classifyConferenceSchedule({
      legacy: { ...v1, enabled: false, running: false, taskState: 1 },
      alternateLegacy: { ...v1, name: 'stock-info-collector', dailyAtLocal: '18:15',
        enabled: true, running: false, taskState: 3 },
      current: v2, expectedUpdateExe, legacyConfigExists: true,
    })).toEqual({ kind: 'blocked', reason: '同時存在可能執行的 v1 與 v2 法說會工作，需先確認唯一寫入者' });
  });

  it('recognizes the older v1 task as a retained legacy writer after it is disabled', () => {
    expect(classifyConferenceSchedule({
      legacy: null,
      alternateLegacy: { ...v1, name: 'stock-info-collector', dailyAtLocal: '18:15',
        enabled: false, running: false, taskState: 1 },
      current: v2, expectedUpdateExe, legacyConfigExists: true,
    })).toEqual({ kind: 'v2', legacyConfigPath: 'F:\\OldReporter\\config\\default.json' });
  });

  it('recognizes both disabled v1 tasks when Windows returns a trailing separator on one working directory', () => {
    expect(classifyConferenceSchedule({
      legacy: { ...v1, enabled: false, running: false, taskState: 1 },
      alternateLegacy: { ...v1, name: 'stock-info-collector', workingDirectory: `${v1.workingDirectory}\\`,
        dailyAtLocal: '18:15', enabled: false, running: false, taskState: 1 },
      current: v2, expectedUpdateExe, legacyConfigExists: true,
    })).toEqual({ kind: 'v2', legacyConfigPath: 'F:\\OldReporter\\config\\default.json' });
  });

  it('creates one v2 task only when both task names are confirmed absent', () => {
    expect(classifyConferenceSchedule({ legacy: null, current: null, expectedUpdateExe, legacyConfigExists: false }))
      .toEqual({ kind: 'create-v2' });
  });

  it('recognizes its own upgrade-stable v2 task and does not recreate it', () => {
    expect(classifyConferenceSchedule({ legacy: null, current: v2, expectedUpdateExe, legacyConfigExists: false }))
      .toEqual({ kind: 'v2' });
  });

  it('does not hide a failed v2 task result or replace that task', () => {
    expect(classifyConferenceSchedule({
      legacy: null, current: { ...v2, lastTaskResult: 0x800710e0 }, expectedUpdateExe, legacyConfigExists: false,
    })).toEqual({ kind: 'v2', lastTaskResult: 0x800710e0 });
  });

  it('explains a disabled v2 task instead of saying the schedule is ready', () => {
    expect(classifyConferenceSchedule({
      legacy: null, current: { ...v2, enabled: false }, expectedUpdateExe, legacyConfigExists: false,
    })).toEqual({
      kind: 'blocked', reason: '偵測到已停用的 v2 法說會工作；目前沒有啟用中的每日法說會排程，未自動修改 Windows 工作',
    });
  });

  it('plans a guarded upgrade for an owned v2 17:45 task with no active legacy writer', () => {
    expect(classifyConferenceSchedule({
      legacy: { ...v1, enabled: false, running: false, taskState: 1 },
      current: { ...v2, dailyAtLocal: '17:45' }, expectedUpdateExe, legacyConfigExists: true,
    })).toEqual({ kind: 'upgrade-v2-time' });
  });

  it('explains a recognized v2 task whose daily time differs from the required 19:00', () => {
    expect(classifyConferenceSchedule({
      legacy: { ...v1, enabled: false, running: false, taskState: 1 },
      current: { ...v2, dailyAtLocal: '18:10' }, expectedUpdateExe, legacyConfigExists: true,
    })).toEqual({
      kind: 'blocked',
      reason: 'v2 法說會工作的每日時間為 18:10，預期 19:00；目前不會執行每日同步，也未自動修改 Windows 工作',
    });
  });

  it('fails closed when inspection is denied, both tasks exist, or a named task is foreign', () => {
    expect(classifyConferenceSchedule({ legacy: 'inaccessible', current: null, expectedUpdateExe, legacyConfigExists: false }).kind).toBe('blocked');
    expect(classifyConferenceSchedule({ legacy: v1, current: v2, expectedUpdateExe, legacyConfigExists: true }).kind).toBe('blocked');
    expect(classifyConferenceSchedule({ legacy: { ...v1, execute: 'C:\\Other\\unknown.exe' }, current: null, expectedUpdateExe, legacyConfigExists: true }).kind).toBe('blocked');
    expect(classifyConferenceSchedule({ legacy: null, current: { ...v2, arguments: '--different' }, expectedUpdateExe, legacyConfigExists: false }).kind).toBe('blocked');
    expect(classifyConferenceSchedule({ legacy: null, current: { ...v2, dailyAtLocal: '18:00' }, expectedUpdateExe, legacyConfigExists: false }).kind).toBe('blocked');
    expect(classifyConferenceSchedule({ legacy: null, current: { ...v2, multipleInstances: 'Parallel' }, expectedUpdateExe, legacyConfigExists: false }).kind).toBe('blocked');
    expect(classifyConferenceSchedule({ legacy: { ...v1, enabled: false }, current: null, expectedUpdateExe, legacyConfigExists: true }).kind).toBe('blocked');
    expect(classifyConferenceSchedule({ legacy: { ...v1, dailyAtLocal: null }, current: null, expectedUpdateExe, legacyConfigExists: true }).kind).toBe('blocked');
    expect(classifyConferenceSchedule({ legacy: v1, current: null, expectedUpdateExe, legacyConfigExists: false }).kind).toBe('blocked');
  });
});
