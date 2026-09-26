import { describe, expect, it, vi } from 'vitest';
import { createMonitoringScheduler } from '../../src/services/scheduler';

const time = (iso: string) => new Date(iso);

describe('monitoring-schedule / scheduler / runs due jobs in Taipei time', () => {
  it('starts a due material check and the once-daily disclosure check after its local time', async () => {
    const calls: string[] = [];
    const scheduler = createMonitoringScheduler({
      now: () => time('2026-09-26T10:31:00.000Z'),
      monitoring: {
        enabled: true, start: '00:00', end: '00:00', intervalMinutes: 60,
        lastCheckedAt: () => time('2026-09-26T09:30:00.000Z'),
        run: async (key) => { calls.push(key); },
      },
      disclosure: {
        enabled: true, runAt: '18:30', lastSuccessfulLocalDate: () => '2026-09-25',
        run: async (key) => { calls.push(key); },
      },
    });

    await expect(scheduler.tick()).resolves.toMatchObject([{ kind: 'material', status: 'started' }, { kind: 'disclosure', status: 'started' }]);
    expect(calls).toHaveLength(2);
    expect(calls[0]).toContain('material:2026-09-26:');
    expect(calls[1]).toBe('disclosure:2026-09-26');
  });
});

describe('monitoring-schedule / catch-up / does not replay every missed interval', () => {
  it('runs one overdue check and does nothing while disabled or outside the configured window', async () => {
    let calls = 0;
    const scheduler = createMonitoringScheduler({
      now: () => time('2026-09-26T01:30:00.000Z'),
      monitoring: {
        enabled: true, start: '10:00', end: '17:00', intervalMinutes: 15,
        lastCheckedAt: () => time('2026-09-25T01:00:00.000Z'),
        run: async () => { calls += 1; },
      },
      disclosure: { enabled: false, runAt: '18:30', lastSuccessfulLocalDate: () => null, run: async () => { calls += 1; } },
    });

    await expect(scheduler.tick()).resolves.toEqual([]);
    expect(calls).toBe(0);
  });
});

describe('monitoring-schedule / overlap / rejects duplicate immediate run while active', () => {
  it('keeps a single active manual run and releases its guard after completion', async () => {
    let finish!: () => void;
    let calls = 0;
    const scheduler = createMonitoringScheduler({
      now: () => time('2026-09-26T01:30:00.000Z'),
      monitoring: {
        enabled: false, start: '00:00', end: '00:00', intervalMinutes: 60, lastCheckedAt: () => null,
        run: async () => { calls += 1; await new Promise<void>((resolve) => { finish = resolve; }); },
      },
      disclosure: { enabled: false, runAt: '18:30', lastSuccessfulLocalDate: () => null, run: async () => undefined },
    });

    const first = scheduler.runNow('material');
    await vi.waitFor(() => expect(calls).toBe(1));
    await expect(scheduler.runNow('material')).resolves.toMatchObject({ status: 'already-running' });
    finish();
    await expect(first).resolves.toMatchObject({ status: 'started' });
    const second = scheduler.runNow('material');
    await vi.waitFor(() => expect(calls).toBe(2));
    finish();
    await expect(second).resolves.toMatchObject({ status: 'started' });
    expect(calls).toBe(2);
  });
});

describe('default-disclosure-monitoring / stale retry / runs once at 19:00 only after stale data', () => {
  it('uses a distinct retry key and never repeats an already-created retry', async () => {
    let calls = 0;
    let retryExists = false;
    const scheduler = createMonitoringScheduler({
      now: () => time('2026-09-26T11:00:00.000Z'),
      monitoring: {
        enabled: false, start: '00:00', end: '00:00', intervalMinutes: 60, lastCheckedAt: () => null,
        run: async () => undefined,
      },
      disclosure: {
        enabled: true, runAt: '18:30', lastSuccessfulLocalDate: () => null,
        lastRunLocalDate: () => '2026-09-26', lastRunStatus: () => 'stale',
        hasRun: (key) => retryExists && key === 'disclosure:2026-09-26:retry-19',
        run: async (key) => { calls += 1; expect(key).toBe('disclosure:2026-09-26:retry-19'); retryExists = true; },
      },
    });

    await expect(scheduler.tick()).resolves.toMatchObject([{ kind: 'disclosure', status: 'started' }]);
    await expect(scheduler.tick()).resolves.toEqual([]);
    expect(calls).toBe(1);
  });
});
