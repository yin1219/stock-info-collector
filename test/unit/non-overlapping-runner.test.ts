import { describe, expect, it } from 'vitest';
import { createNonOverlappingRunner } from '../../src/services/job-runner';

describe('monitoring-schedule / overlap / rejects duplicate immediate run while active', () => {
  it('does not start a second run until the first settles, then allows another run', async () => {
    let finishFirst!: (value: string) => void;
    let calls = 0;
    const runner = createNonOverlappingRunner(() => {
      calls += 1;
      if (calls === 1) return new Promise<string>((resolve) => { finishFirst = resolve; });
      return Promise.resolve('second complete');
    });

    const firstRun = runner.run();
    expect(runner.isRunning()).toBe(true);
    await expect(runner.run()).resolves.toEqual({ started: false, reason: 'already-running' });
    expect(calls).toBe(1);
    finishFirst('first complete');
    await expect(firstRun).resolves.toEqual({ started: true, value: 'first complete' });
    expect(runner.isRunning()).toBe(false);
    await expect(runner.run()).resolves.toEqual({ started: true, value: 'second complete' });
  });
});

describe('monitoring-schedule / overlap / releases the lock after an execution error', () => {
  it('propagates a source error and does not leave the job permanently active', async () => {
    let shouldFail = true;
    const runner = createNonOverlappingRunner(async () => {
      if (shouldFail) throw new Error('provider failed');
      return 'recovered';
    });

    await expect(runner.run()).rejects.toThrow('provider failed');
    expect(runner.isRunning()).toBe(false);
    shouldFail = false;
    await expect(runner.run()).resolves.toEqual({ started: true, value: 'recovered' });
  });
});
