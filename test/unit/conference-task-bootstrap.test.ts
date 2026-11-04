import { describe, expect, it, vi } from 'vitest';
import { bootstrapConferenceTask } from '../../src/main/conference-task-bootstrap';
import { LEGACY_CONFERENCE_TASK_NAME, V2_CONFERENCE_TASK_NAME, V2_CONFERENCE_TASK_ARGUMENTS } from '../../src/main/conference-task-plan';

const expectedUpdateExe = 'C:\\Users\\Test\\AppData\\Local\\StockReporterAssistant\\Update.exe';
const legacy = {
  name: LEGACY_CONFERENCE_TASK_NAME, execute: 'F:\\OldReporter\\stock-info-collector.exe',
  workingDirectory: 'F:\\OldReporter', arguments: '', enabled: true, dailyAtLocal: '19:00',
};
const current = {
  name: V2_CONFERENCE_TASK_NAME, execute: expectedUpdateExe,
  workingDirectory: 'C:\\Users\\Test\\AppData\\Local\\StockReporterAssistant',
  arguments: V2_CONFERENCE_TASK_ARGUMENTS, enabled: true,
  dailyAtLocal: '19:00', multipleInstances: 'IgnoreNew',
};

describe('windows-distribution / conference task bootstrap / preserves one writer', () => {
  it('preserves a recognized v1 task without registering or changing it', async () => {
    const registerIfAbsent = vi.fn();
    const outcome = await bootstrapConferenceTask({
      inspect: async (name) => name === LEGACY_CONFERENCE_TASK_NAME ? legacy : null,
      legacyConfigExists: async () => true,
      registerIfAbsent, expectedUpdateExe, setupReady: false,
    });
    expect(outcome).toEqual({ kind: 'legacy', configPath: 'F:\\OldReporter\\config\\default.json' });
    expect(registerIfAbsent).not.toHaveBeenCalled();
  });

  it('registers a v2 daily job only when both names are confirmed absent and verifies it', async () => {
    let created = false;
    const registerIfAbsent = vi.fn(async () => { created = true; });
    const outcome = await bootstrapConferenceTask({
      inspect: async (name) => name === V2_CONFERENCE_TASK_NAME && created ? current : null,
      legacyConfigExists: async () => false,
      registerIfAbsent, expectedUpdateExe, setupReady: true,
    });
    expect(outcome).toEqual({ kind: 'v2' });
    expect(registerIfAbsent).toHaveBeenCalledTimes(1);
    expect(registerIfAbsent).toHaveBeenCalledWith({
      name: V2_CONFERENCE_TASK_NAME, execute: expectedUpdateExe,
      workingDirectory: current.workingDirectory,
      arguments: V2_CONFERENCE_TASK_ARGUMENTS,
      localTime: '19:00', multipleInstances: 'IgnoreNew',
    });
  });

  it('never registers when inspection is denied, names collide, or a v2 task already exists', async () => {
    for (const [first, second] of [
      ['inaccessible', null], [legacy, current], [null, { ...current, arguments: '--foreign' }], [null, current],
    ] as const) {
      const registerIfAbsent = vi.fn();
      const result = await bootstrapConferenceTask({
        inspect: async (name) => name === LEGACY_CONFERENCE_TASK_NAME ? first
          : name === V2_CONFERENCE_TASK_NAME ? second : null,
        legacyConfigExists: async () => true,
        registerIfAbsent, expectedUpdateExe, setupReady: true,
      });
      expect(registerIfAbsent).not.toHaveBeenCalled();
      expect(result.kind).not.toBe('create-v2');
    }
  });

  it('inspects the older v1 task name and refuses a second Calendar writer', async () => {
    const inspected: string[] = [];
    const result = await bootstrapConferenceTask({
      inspect: async (name) => {
        inspected.push(name);
        if (name === 'stock-info-collector') return {
          ...legacy, name, dailyAtLocal: '18:15', enabled: true, running: false, taskState: 3,
        };
        return name === V2_CONFERENCE_TASK_NAME ? current : null;
      },
      legacyConfigExists: async () => true, registerIfAbsent: vi.fn(),
      expectedUpdateExe, setupReady: true,
    });
    expect(inspected).toContain('stock-info-collector');
    expect(result.kind).toBe('blocked');
  });

  it('reports registration or verification failure without claiming a working schedule', async () => {
    const failed = await bootstrapConferenceTask({
      inspect: async () => null, legacyConfigExists: async () => false,
      registerIfAbsent: async () => { throw new Error('access denied'); }, expectedUpdateExe, setupReady: true,
    });
    expect(failed.kind).toBe('blocked');

    const unverified = await bootstrapConferenceTask({
      inspect: async () => null, legacyConfigExists: async () => false,
      registerIfAbsent: async () => undefined, expectedUpdateExe, setupReady: true,
    });
    expect(unverified.kind).toBe('blocked');
  });

  it('waits for OAuth and an active watchlist before creating a new v2 task', async () => {
    const registerIfAbsent = vi.fn();
    const outcome = await bootstrapConferenceTask({
      inspect: async () => null, legacyConfigExists: async () => false,
      registerIfAbsent, expectedUpdateExe, setupReady: false,
    });
    expect(outcome.kind).toBe('blocked');
    expect(registerIfAbsent).not.toHaveBeenCalled();
  });

  it('switches only an explicitly requested disabled and stopped v1 task, then verifies v2 ownership', async () => {
    let created = false;
    const registerIfAbsent = vi.fn(async () => { created = true; });
    const outcome = await bootstrapConferenceTask({
      inspect: async (name) => name === LEGACY_CONFERENCE_TASK_NAME
        ? { ...legacy, enabled: false, running: false, taskState: 1 }
        : name === V2_CONFERENCE_TASK_NAME && created ? current : null,
      legacyConfigExists: async () => true,
      registerIfAbsent, expectedUpdateExe, setupReady: true, switchDisabledLegacy: true,
    });
    expect(outcome).toEqual({ kind: 'v2', legacyConfigPath: 'F:\\OldReporter\\config\\default.json' });
    expect(registerIfAbsent).toHaveBeenCalledWith(expect.objectContaining({ allowDisabledLegacy: true }));
  });

  it('does not switch without explicit request, setup, or confirmed stopped legacy state', async () => {
    for (const [oldTask, ready, switchDisabledLegacy] of [
      [{ ...legacy, enabled: false, running: false, taskState: 1 }, true, false],
      [{ ...legacy, enabled: false, running: false, taskState: 1 }, false, true],
      [{ ...legacy, enabled: false, running: true, taskState: 4 }, true, true],
      [legacy, true, true],
    ] as const) {
      const registerIfAbsent = vi.fn();
      const outcome = await bootstrapConferenceTask({
        inspect: async (name) => name === LEGACY_CONFERENCE_TASK_NAME ? oldTask : null,
        legacyConfigExists: async () => true,
        registerIfAbsent, expectedUpdateExe, setupReady: ready, switchDisabledLegacy,
      });
      expect(registerIfAbsent).not.toHaveBeenCalled();
      expect(outcome.kind).not.toBe('v2');
    }
  });

  it('does not turn an explicit legacy switch into first-install registration when v1 disappears', async () => {
    const registerIfAbsent = vi.fn();
    const result = await bootstrapConferenceTask({
      inspect: async () => null, legacyConfigExists: async () => false,
      registerIfAbsent, expectedUpdateExe, setupReady: true, switchDisabledLegacy: true,
    });
    expect(result.kind).toBe('blocked');
    expect(registerIfAbsent).not.toHaveBeenCalled();
  });

  it('migrates an owned 17:45 v2 task once and verifies the new 19:00 writer', async () => {
    let migrated = false;
    const migrateOwnedTime = vi.fn(async () => { migrated = true; });
    const result = await bootstrapConferenceTask({
      inspect: async (name) => name === LEGACY_CONFERENCE_TASK_NAME
        ? { ...legacy, enabled: false, running: false, taskState: 1 }
        : name === V2_CONFERENCE_TASK_NAME
          ? { ...current, dailyAtLocal: migrated ? '19:00' : '17:45' } : null,
      legacyConfigExists: async () => true, registerIfAbsent: vi.fn(),
      migrateOwnedTime, expectedUpdateExe, setupReady: true,
    });
    expect(result).toEqual({ kind: 'v2', legacyConfigPath: 'F:\\OldReporter\\config\\default.json' });
    expect(migrateOwnedTime).toHaveBeenCalledExactlyOnceWith(expectedUpdateExe);
  });

  it('does not migrate an unknown task, an active v1 writer, or an unverified result', async () => {
    const migrateOwnedTime = vi.fn();
    for (const [oldTask, newTask] of [
      [legacy, { ...current, dailyAtLocal: '17:45' }],
      [null, { ...current, dailyAtLocal: '18:00' }],
      [null, { ...current, dailyAtLocal: '17:45', arguments: '--foreign' }],
    ] as const) {
      const result = await bootstrapConferenceTask({
        inspect: async (name) => name === LEGACY_CONFERENCE_TASK_NAME ? oldTask
          : name === V2_CONFERENCE_TASK_NAME ? newTask : null,
        legacyConfigExists: async () => true, registerIfAbsent: vi.fn(), migrateOwnedTime,
        expectedUpdateExe, setupReady: true,
      });
      expect(result.kind).toBe('blocked');
    }
    expect(migrateOwnedTime).not.toHaveBeenCalled();

    const unverified = await bootstrapConferenceTask({
      inspect: async (name) => name === V2_CONFERENCE_TASK_NAME ? { ...current, dailyAtLocal: '17:45' } : null,
      legacyConfigExists: async () => false, registerIfAbsent: vi.fn(),
      migrateOwnedTime: vi.fn(), expectedUpdateExe, setupReady: true,
    });
    expect(unverified.kind).toBe('blocked');
  });
});
