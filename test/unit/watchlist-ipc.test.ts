import { describe, expect, it } from 'vitest';
import { registerWatchlistIpc, type IpcMainPort } from '../../src/main/watchlist-ipc';

describe('watchlist-management / ipc / exposes validated list and edit operations', () => {
  it('registers only explicit watchlist operations and forwards typed data', async () => {
    const handlers = new Map<string, (...args: unknown[]) => unknown>();
    const ipcMain: IpcMainPort = { handle(channel, handler) { handlers.set(channel, handler); } };
    const service = {
      list: (filter: unknown) => ({ filter }),
      add: (input: unknown) => ({ input }),
      update: (input: unknown) => ({ input }),
      setActive: (input: unknown) => ({ input }),
      remove: (input: unknown) => ({ input }),
      previewImport: (input: unknown) => ({ input }),
      applyImport: (input: unknown) => ({ input }),
    };

    registerWatchlistIpc(ipcMain, service, (event) => event.senderId === 'main-window');

    expect([...handlers.keys()]).toEqual([
      'watchlist:list', 'watchlist:add', 'watchlist:update', 'watchlist:set-active',
      'watchlist:remove', 'watchlist:import-preview', 'watchlist:import-apply',
    ]);
    await expect(handlers.get('watchlist:add')!({ senderId: 'main-window' }, { market: 'TWSE', stockCode: '2330' }))
      .resolves.toEqual({ input: { market: 'TWSE', stockCode: '2330' } });
  });
});

describe('security architecture / ipc rejects untrusted renderer senders', () => {
  it('rejects an IPC caller not authorized as the application window', async () => {
    const handlers = new Map<string, (...args: unknown[]) => unknown>();
    const ipcMain: IpcMainPort = { handle(channel, handler) { handlers.set(channel, handler); } };
    const service = Object.fromEntries([
      'list', 'add', 'update', 'setActive', 'remove', 'previewImport', 'applyImport',
    ].map((key) => [key, () => 'should not run']));
    registerWatchlistIpc(ipcMain, service, () => false);

    await expect(handlers.get('watchlist:list')!({ senderId: 'foreign' }, {})).rejects.toThrow(/不受信任/);
  });
});
