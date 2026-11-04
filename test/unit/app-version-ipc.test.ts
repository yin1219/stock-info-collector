import { describe, expect, it, vi } from 'vitest';
import { registerAppVersionIpc } from '../../src/main/app-version-ipc';

describe('desktop-app-lifecycle / installed version', () => {
  it('returns the running app version only to the trusted renderer and unregisters on quit', async () => {
    const handlers = new Map<string, (event: unknown) => unknown>();
    const removeHandler = vi.fn((channel: string) => handlers.delete(channel));
    const getVersion = vi.fn(() => '1.0.13');
    const trustedEvent = {};
    const unregister = registerAppVersionIpc({
      handle(channel, listener) { handlers.set(channel, listener); }, removeHandler,
    }, { getVersion }, (event) => event === trustedEvent);

    const handler = handlers.get('desktop:get-app-version');
    expect(handler).toBeDefined();
    expect(() => handler?.({})).toThrow('拒絕不受信任的 renderer IPC 呼叫');
    expect(getVersion).not.toHaveBeenCalled();
    expect(await handler?.(trustedEvent)).toBe('1.0.13');
    unregister();
    expect(removeHandler).toHaveBeenCalledWith('desktop:get-app-version');
  });
});
