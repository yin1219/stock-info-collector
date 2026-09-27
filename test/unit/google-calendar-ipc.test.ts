import { describe, expect, it, vi } from 'vitest';
import { registerGoogleCalendarIpc } from '../../src/main/google-calendar-ipc';

describe('conference-calendar-sync / Google Calendar IPC', () => {
  it('restricts calendar status, credential selection and writes to the trusted main frame', async () => {
    const handlers = new Map<string, (event: unknown) => unknown>();
    const controller = {
      status: vi.fn(async () => ({ status: 'connected' })),
      list: vi.fn(() => []),
      importCredentials: vi.fn(async () => true),
      migrateLegacyToken: vi.fn(async () => true),
      connect: vi.fn(async () => ({ status: 'connected' })),
      disconnect: vi.fn(async () => ({ status: 'disconnected' })),
      sync: vi.fn(async () => ({ status: 'complete', results: [] })),
    };
    const unregister = registerGoogleCalendarIpc({
      handle: (channel, listener) => handlers.set(channel, listener as (event: unknown) => unknown),
      removeHandler: (channel) => { handlers.delete(channel); },
    }, controller, (event) => event === 'trusted');

    expect(() => handlers.get('calendar:get-status')!('untrusted')).toThrow(/不受信任/);
    await expect(handlers.get('calendar:get-status')!('trusted')).resolves.toEqual({ status: 'connected' });
    await expect(handlers.get('calendar:import-credentials')!('trusted')).resolves.toBe(true);
    await expect(handlers.get('calendar:migrate-legacy-token')!('trusted')).resolves.toBe(true);
    await expect(handlers.get('calendar:connect')!('trusted')).resolves.toEqual({ status: 'connected' });
    await expect(handlers.get('calendar:sync')!('trusted')).resolves.toEqual({ status: 'complete', results: [] });
    await expect(handlers.get('calendar:disconnect')!('trusted')).resolves.toEqual({ status: 'disconnected' });
    expect(controller.sync).toHaveBeenCalledOnce();
    unregister();
    expect(handlers.size).toBe(0);
  });

  it('records an OAuth connection failure by operation while preserving the UI error', async () => {
    const handlers = new Map<string, (event: unknown) => unknown>();
    const failure = new Error('Google OAuth authorization timed out');
    const onFailure = vi.fn();
    registerGoogleCalendarIpc({
      handle: (channel, listener) => handlers.set(channel, listener as (event: unknown) => unknown),
    }, {
      status: () => undefined,
      list: () => [],
      importCredentials: () => undefined,
      migrateLegacyToken: () => undefined,
      connect: () => Promise.reject(failure),
      disconnect: () => undefined,
      sync: () => undefined,
    }, () => true, onFailure);

    await expect(handlers.get('calendar:connect')!('trusted')).rejects.toBe(failure);
    expect(onFailure).toHaveBeenCalledExactlyOnceWith('calendar:connect', failure);
  });
});
