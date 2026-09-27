import { describe, expect, it, vi } from 'vitest';
import { createScheduleSettingsController } from '../../src/services/schedule-settings';

function fixture() {
  const values = new Map<string, unknown>();
  const repositories = {
    settings: { get: <T>(key: string) => values.get(key) as T | undefined, set: (key: string, value: unknown) => { values.set(key, value); } },
    jobRuns: { list: vi.fn(() => []) },
  } as never;
  const scheduler = {
    updateSettings: vi.fn(),
    runNow: vi.fn(async (kind: 'material' | 'disclosure') => ({ kind, status: 'started' as const })),
    isRunning: vi.fn(() => false),
  };
  return { controller: createScheduleSettingsController({ repositories, scheduler, now: () => new Date('2026-09-26T03:00:00.000Z') }), scheduler, values, repositories };
}

describe('monitoring-schedule / settings controller', () => {
  it('returns defaults and persists valid settings into the live scheduler', () => {
    const { controller, scheduler, values } = fixture();
    expect(controller.getSettings()).toEqual({
      monitoring: { enabled: true, start: '00:00', end: '00:00', intervalMinutes: 60 },
      disclosure: { enabled: true, runAt: '18:30' },
      notifyEmptyDefaultDisclosures: true,
    });
    const updated = {
      monitoring: { enabled: false, start: '08:00', end: '23:00', intervalMinutes: 30 },
      disclosure: { enabled: true, runAt: '19:00' },
      notifyEmptyDefaultDisclosures: true,
    };
    expect(controller.saveSettings(updated)).toEqual(updated);
    expect(values.get('schedule')).toEqual(updated);
    expect(scheduler.updateSettings).toHaveBeenCalledWith({ monitoring: updated.monitoring, disclosure: updated.disclosure });
  });

  it('defaults empty-disclosure notifications to enabled and persists the user preference', () => {
    const { controller, values } = fixture();
    expect(controller.getSettings().notifyEmptyDefaultDisclosures).toBe(true);
    const updated = { ...controller.getSettings(), notifyEmptyDefaultDisclosures: false };

    expect(controller.saveSettings(updated).notifyEmptyDefaultDisclosures).toBe(false);
    expect(values.get('notifyEmptyDefaultDisclosures')).toBe(false);
    expect(controller.getSettings().notifyEmptyDefaultDisclosures).toBe(false);
  });

  it('rejects unsupported intervals and malformed times without saving', () => {
    const { controller, values } = fixture();
    expect(() => controller.saveSettings({ monitoring: { enabled: true, start: '08:00', end: '23:00', intervalMinutes: 10 }, disclosure: { enabled: true, runAt: '18:30' } })).toThrow(/頻率/);
    expect(() => controller.saveSettings({ monitoring: { enabled: true, start: '8am', end: '23:00', intervalMinutes: 60 }, disclosure: { enabled: true, runAt: '18:30' } })).toThrow(/時間/);
    expect(values.has('schedule')).toBe(false);
  });

  it('reports the latest monitor and disclosure execution states and offers a manual run', async () => {
    const { controller, scheduler, repositories } = fixture();
    vi.mocked(repositories.jobRuns.list).mockImplementation(((kind: string) => kind === 'material-event-monitoring'
      ? [{ status: 'failed', finishedAt: '2026-09-26T02:50:00.000Z', errorMessage: 'MOPS unavailable' }]
      : [{ status: 'complete', finishedAt: '2026-09-26T02:00:00.000Z', errorMessage: null }]) as never);
    expect(controller.getStatus()).toMatchObject({
      material: { status: 'failed', errorMessage: 'MOPS unavailable' },
      disclosure: { status: 'complete' },
    });
    await expect(controller.runNow('material')).resolves.toMatchObject({ status: 'started' });
    expect(scheduler.runNow).toHaveBeenCalledWith('material');
  });

  it('surfaces already-running instead of starting a duplicate', async () => {
    const { controller, scheduler } = fixture();
    scheduler.runNow.mockResolvedValue({ kind: 'material', status: 'already-running' });
    await expect(controller.runNow('material')).resolves.toMatchObject({ status: 'already-running' });
  });

  it('merges a partial persisted value with defaults and calculates the next local disclosure date', () => {
    const { controller, values } = fixture();
    values.set('schedule', { disclosure: { enabled: true, runAt: '18:30' } });
    expect(controller.getSettings()).toEqual({
      monitoring: { enabled: true, start: '00:00', end: '00:00', intervalMinutes: 60 },
      disclosure: { enabled: true, runAt: '18:30' },
      notifyEmptyDefaultDisclosures: true,
    });
    const late = createScheduleSettingsController({
      repositories: {
        settings: { get: (key: string) => key === 'schedule' ? ({ monitoring: { enabled: false, start: '00:00', end: '00:00', intervalMinutes: 60 }, disclosure: { enabled: false, runAt: '18:30' } }) : undefined, set: vi.fn() },
        jobRuns: { list: vi.fn(() => []) },
      } as never,
      now: () => new Date('2026-09-26T15:00:00.000Z'),
    });
    expect(late.getStatus().disclosure.nextRunAt).toBeNull();
    expect(late.getStatus().material.nextRunAt).toBeNull();
  });

  it('reports live scheduler activity and rejects manual work when the scheduler is unavailable', async () => {
    const { controller } = fixture();
    const unavailable = createScheduleSettingsController({
      repositories: { settings: { get: vi.fn(), set: vi.fn() }, jobRuns: { list: vi.fn(() => []) } } as never,
      now: () => new Date('2026-09-26T03:00:00.000Z'),
    });
    expect(() => unavailable.runNow('material')).toThrow(/尚未啟動/);
    expect(unavailable.getStatus().available).toBe(false);
    const active = createScheduleSettingsController({
      repositories: { settings: { get: vi.fn(), set: vi.fn() }, jobRuns: { list: vi.fn(() => []) } } as never,
      scheduler: { updateSettings: vi.fn(), runNow: vi.fn(), isRunning: (kind) => kind === 'material' },
      now: () => new Date('2026-09-26T03:00:00.000Z'),
    });
    expect(active.getStatus().material.status).toBe('active');
    expect(active.getStatus().available).toBe(true);
    expect(active.getStatus().disclosure.status).toBe('idle');
    await expect(controller.runNow('disclosure')).resolves.toMatchObject({ status: 'started' });
  });
});
