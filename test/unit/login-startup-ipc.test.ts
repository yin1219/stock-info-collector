import { describe, expect, it, vi } from 'vitest';
import { registerLoginStartupIpc } from '../../src/main/login-startup-ipc';

describe('desktop-app-lifecycle / login startup IPC', () => {
  it('validates trusted login settings before delegating to the OS controller', () => {
    const handlers = new Map<string, (event: unknown, payload?: unknown) => unknown>();
    const ipc = { handle: vi.fn((channel: string, handler: (event: unknown, payload?: unknown) => unknown) => handlers.set(channel, handler)), removeHandler: vi.fn() };
    const controller = { getSettings: vi.fn(() => ({ enabled: false, startHidden: false })), saveSettings: vi.fn((value) => value) };
    const unregister = registerLoginStartupIpc(ipc, controller, (event) => event === 'trusted');
    expect(handlers.get('desktop:get-login-startup')!('trusted')).toEqual({ enabled: false, startHidden: false });
    expect(handlers.get('desktop:save-login-startup')!('trusted', { enabled: true, startHidden: true }))
      .toEqual({ enabled: true, startHidden: true });
    expect(() => handlers.get('desktop:save-login-startup')!('trusted', { enabled: 'yes', startHidden: false }))
      .toThrow('登入啟動設定格式無效');
    expect(() => handlers.get('desktop:get-login-startup')!('untrusted')).toThrow('拒絕不受信任');
    unregister();
    expect(ipc.removeHandler).toHaveBeenCalledTimes(2);
  });
});
