import { describe, expect, it, vi } from 'vitest';
import { registerConferenceTaskIpc } from '../../src/main/conference-task-ipc';

describe('windows-distribution / conference task IPC / trusted first-run controls', () => {
  it('exposes status, explicit setup, legacy review, and confirmed StockNumbers apply only', async () => {
    const handlers = new Map<string, (event: unknown, payload?: unknown) => Promise<unknown>>();
    const service = {
      status: vi.fn(async () => ({ kind: 'legacy' })),
      ensure: vi.fn(async () => ({ kind: 'v2' })),
      switchFromDisabledLegacy: vi.fn(async () => ({ kind: 'v2' })),
      reviewLegacy: vi.fn(async () => ({ hash: 'a'.repeat(64) })),
      confirmLegacy: vi.fn(async (hash: string) => ({ hash })),
    };
    const unregister = registerConferenceTaskIpc({
      handle: (channel, listener) => { handlers.set(channel, listener); },
      removeHandler: (channel) => { handlers.delete(channel); },
    }, service, (event) => event === 'trusted');

    expect([...handlers.keys()]).toEqual([
      'conference-task:status', 'conference-task:ensure', 'conference-task:switch-disabled-legacy', 'conference-task:review-legacy', 'conference-task:confirm-legacy',
    ]);
    await expect(handlers.get('conference-task:status')!('trusted')).resolves.toEqual({ kind: 'legacy' });
    await expect(handlers.get('conference-task:confirm-legacy')!('trusted', { hash: 'a'.repeat(64) }))
      .resolves.toEqual({ hash: 'a'.repeat(64) });
    await expect(handlers.get('conference-task:confirm-legacy')!('trusted', { hash: 'bad' })).rejects.toThrow(/核對/);
    await expect(handlers.get('conference-task:ensure')!('foreign')).rejects.toThrow(/不受信任/);
    await expect(handlers.get('conference-task:switch-disabled-legacy')!('trusted', {})).rejects.toThrow(/確認/);
    await expect(handlers.get('conference-task:switch-disabled-legacy')!('trusted', { confirmation: 'SWITCH_TO_V2_DAILY' }))
      .resolves.toEqual({ kind: 'v2' });
    expect(service.ensure).not.toHaveBeenCalled();
    unregister();
    expect(handlers.size).toBe(0);
  });
});
