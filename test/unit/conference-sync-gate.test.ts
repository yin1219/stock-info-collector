import { describe, expect, it, vi } from 'vitest';
import { createConferenceSyncGate } from '../../src/main/conference-sync-gate';

describe('conference-calendar-sync / single writer / manual and scheduled entry share a gate', () => {
  it('refuses v2 Calendar writes while an existing v1 daily writer owns the task', async () => {
    const runSync = vi.fn();
    const gate = createConferenceSyncGate({ taskOwner: async () => 'legacy', runSync });
    await expect(gate.sync()).rejects.toThrow(/v1.*工作/);
    expect(runSync).not.toHaveBeenCalled();
  });

  it('fails closed when task inspection cannot identify an owner', async () => {
    const runSync = vi.fn();
    const gate = createConferenceSyncGate({ taskOwner: async () => 'blocked', runSync });
    await expect(gate.sync()).rejects.toThrow(/無法確認/);
    expect(runSync).not.toHaveBeenCalled();
  });

  it('allows a one-at-a-time manual Calendar acceptance with a verified stopped and disabled v1 task', async () => {
    const runSync = vi.fn(async () => 'accepted');
    const gate = createConferenceSyncGate({ taskOwner: async () => 'legacy-disabled', runSync });
    await expect(gate.sync()).resolves.toBe('accepted');
    expect(runSync).toHaveBeenCalledOnce();
  });

  it('runs one v2 sync at a time and releases the gate after failure', async () => {
    let resolve!: (value: string) => void;
    const runSync = vi.fn().mockImplementationOnce(() => new Promise<string>((done) => { resolve = done; }))
      .mockRejectedValueOnce(new Error('temporary failure')).mockResolvedValueOnce('recovered');
    const gate = createConferenceSyncGate({ taskOwner: async () => 'v2', runSync });
    const first = gate.sync();
    await vi.waitFor(() => expect(runSync).toHaveBeenCalledOnce());
    await expect(gate.sync()).rejects.toThrow(/正在執行/);
    resolve('done');
    await expect(first).resolves.toBe('done');
    await expect(gate.sync()).rejects.toThrow(/temporary failure/);
    await expect(gate.sync()).resolves.toBe('recovered');
  });
});
