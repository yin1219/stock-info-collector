import { describe, expect, it, vi } from 'vitest';
import { registerScheduleIpc } from '../../src/main/schedule-ipc';

function fixture() {
  const handlers = new Map<string, (event: unknown, payload?: unknown) => unknown>();
  const service = {
    getSettings: vi.fn(() => ({ monitoring: {}, disclosure: {} })),
    saveSettings: vi.fn((value) => value),
    getStatus: vi.fn(() => ({ material: { status: 'idle' } })),
    runNow: vi.fn(async (kind) => ({ kind, status: 'started' })),
  };
  const unregister = registerScheduleIpc({
    handle: (channel, listener) => handlers.set(channel, listener),
    removeHandler: (channel) => { handlers.delete(channel); },
  }, service as never, (event) => event === 'trusted');
  return { handlers, service, unregister };
}

describe('monitoring-schedule / IPC', () => {
  it('requires the trusted main frame and exposes settings, status and manual runs', async () => {
    const { handlers, service, unregister } = fixture();
    expect(() => handlers.get('schedule:get-settings')!('untrusted')).toThrow(/不受信任/);
    expect(handlers.get('schedule:get-settings')!('trusted')).toEqual({ monitoring: {}, disclosure: {} });
    expect(handlers.get('schedule:get-status')!('trusted')).toEqual({ material: { status: 'idle' } });
    await expect(handlers.get('schedule:run-now')!('trusted', { kind: 'material' })).resolves.toEqual({ kind: 'material', status: 'started' });
    expect(service.runNow).toHaveBeenCalledWith('material');
    unregister();
    expect(handlers.size).toBe(0);
  });

  it('rejects an invalid manual run kind', () => {
    const { handlers, service } = fixture();
    expect(() => handlers.get('schedule:run-now')!('trusted', { kind: 'arbitrary' })).toThrow(/工作種類/);
    expect(service.runNow).not.toHaveBeenCalled();
  });
});
