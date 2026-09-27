import { describe, expect, it, vi } from 'vitest';
import { startSchedulerLifecycle } from '../../src/main/scheduler-bootstrap';

describe('scheduler lifecycle bootstrap', () => {
  it('does not create a live scheduler in isolated test userData mode', () => {
    const createScheduler = vi.fn();
    const lifecycle = startSchedulerLifecycle({ isolated: true, createScheduler, powerMonitor: { on: vi.fn(), removeListener: vi.fn() }, setInterval: vi.fn(), clearInterval: vi.fn() });
    expect(createScheduler).not.toHaveBeenCalled();
    expect(lifecycle).toBeUndefined();
  });

  it('allows an explicitly opted-in isolated scheduler to run manually without startup, polling or resume work', async () => {
    const tick = vi.fn().mockResolvedValue([]);
    const runNow = vi.fn().mockResolvedValue({ status: 'started' });
    const on = vi.fn();
    const interval = vi.fn();
    const lifecycle = startSchedulerLifecycle({
      isolated: true,
      manualOnly: true,
      createScheduler: () => ({ tick, catchUp: tick, runNow }),
      powerMonitor: { on, removeListener: vi.fn() },
      setInterval: interval,
      clearInterval: vi.fn(),
    });
    expect(lifecycle?.scheduler).toBeDefined();
    expect(tick).not.toHaveBeenCalled();
    expect(interval).not.toHaveBeenCalled();
    expect(on).not.toHaveBeenCalled();
    await expect(lifecycle!.scheduler.runNow?.('material')).resolves.toEqual({ status: 'started' });
    lifecycle?.stop();
  });

  it('checks on startup and resume, polls, and removes timer/listener on stop', async () => {
    const tick = vi.fn().mockResolvedValue([]);
    const on = vi.fn();
    const removeListener = vi.fn();
    const interval = vi.fn(() => 17);
    const clearInterval = vi.fn();
    const lifecycle = startSchedulerLifecycle({
      isolated: false, createScheduler: () => ({ tick, catchUp: tick }),
      powerMonitor: { on, removeListener }, setInterval: interval, clearInterval,
    })!;
    expect(tick).toHaveBeenCalledTimes(1);
    expect(interval).toHaveBeenCalledWith(expect.any(Function), 60_000);
    expect(on).toHaveBeenCalledWith('resume', expect.any(Function));
    await on.mock.calls[0][1]();
    expect(tick).toHaveBeenCalledTimes(2);
    lifecycle.stop();
    expect(clearInterval).toHaveBeenCalledWith(17);
    expect(removeListener).toHaveBeenCalledWith('resume', on.mock.calls[0][1]);
  });
});
