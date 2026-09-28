import { describe, expect, it, vi } from 'vitest';
import { registerTestNotificationIpc } from '../../src/main/test-notification-ipc';

describe('notification-delivery / test notification IPC', () => {
  it('rejects untrusted requests, delegates trusted requests, and unregisters the handler', async () => {
    const handlers = new Map<string, (event: unknown) => unknown>();
    const send = vi.fn(async () => ({ status: 'requested' as const }));
    const schedule = vi.fn(() => ({ status: 'scheduled' as const, scheduledAt: '2026-09-28T00:01:00.000Z' }));
    const onFailure = vi.fn();
    const unregister = registerTestNotificationIpc({
      handle: (channel, handler) => handlers.set(channel, handler),
      removeHandler: (channel) => handlers.delete(channel),
    }, { send, schedule }, (event) => event === 'trusted', onFailure);

    const handler = handlers.get('notification:send-test')!;
    expect(() => handler('untrusted')).toThrow('拒絕不受信任');
    expect(send).not.toHaveBeenCalled();
    await expect(handler('trusted')).resolves.toEqual({ status: 'requested' });
    expect(send).toHaveBeenCalledTimes(1);
    const delayed = handlers.get('notification:schedule-test')!;
    expect(() => delayed('untrusted')).toThrow('拒絕不受信任');
    await expect(delayed('trusted')).resolves.toEqual({ status: 'scheduled', scheduledAt: '2026-09-28T00:01:00.000Z' });
    expect(schedule).toHaveBeenCalledTimes(1);
    expect(onFailure).not.toHaveBeenCalled();
    unregister();
    expect(handlers.has('notification:send-test')).toBe(false);
    expect(handlers.has('notification:schedule-test')).toBe(false);
  });

  it('records a native failure without swallowing it', async () => {
    const handlers = new Map<string, (event: unknown) => unknown>();
    const onFailure = vi.fn();
    registerTestNotificationIpc({ handle: (channel, handler) => handlers.set(channel, handler) },
      { send: async () => { throw new Error('native unavailable'); }, schedule: () => ({ status: 'suppressed' }) }, () => true, onFailure);

    await expect(handlers.get('notification:send-test')!({})).rejects.toThrow('native unavailable');
    expect(onFailure).toHaveBeenCalledWith(expect.any(Error));
  });
});
