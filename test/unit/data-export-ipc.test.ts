import { describe, expect, it, vi } from 'vitest';
import { registerDataExportIpc } from '../../src/main/data-export-ipc';

describe('local-data-management / export IPC', () => {
  it('allows only the trusted renderer to request a native export and unregisters the handler', async () => {
    const handlers = new Map<string, (event: unknown) => unknown>();
    const ipc = {
      handle: vi.fn((channel: string, handler: (event: unknown) => unknown) => handlers.set(channel, handler)),
      removeHandler: vi.fn((channel: string) => handlers.delete(channel)),
    };
    const controller = { exportUserData: vi.fn(async () => ({ status: 'exported' as const, destination: 'C:/safe/export.json' })) };
    const unregister = registerDataExportIpc(ipc, controller, (event) => event === 'trusted');

    await expect(handlers.get('data:export')!('trusted')).resolves.toEqual({ status: 'exported', destination: 'C:/safe/export.json' });
    expect(() => handlers.get('data:export')!('untrusted')).toThrow('拒絕不受信任的 renderer IPC 呼叫');
    expect(controller.exportUserData).toHaveBeenCalledOnce();
    unregister();
    expect(ipc.removeHandler).toHaveBeenCalledWith('data:export');
  });
});
