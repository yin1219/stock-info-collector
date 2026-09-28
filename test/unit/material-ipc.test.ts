import { describe, expect, it } from 'vitest';
import { registerMaterialIpc } from '../../src/main/material-ipc';

function fakeIpc() {
  const handlers = new Map<string, (event: unknown, payload?: unknown) => unknown>();
  return { handlers, handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown) { handlers.set(channel, listener); }, removeHandler(channel: string) { handlers.delete(channel); } };
}

describe('material-event-monitoring / IPC', () => {
  it('limits operations to local listing, details, read-state and confirmed soft-deletion methods', () => {
    const ipc = fakeIpc();
    const calls: unknown[] = [];
    const dispose = registerMaterialIpc(ipc, {
      list(filter) { calls.push(['list', filter]); return []; },
      detail(id) { calls.push(['detail', id]); return { id }; },
      markRead(id) { calls.push(['read', id]); return { id }; },
      softDelete(id) { calls.push(['delete', id]); return { id }; },
      monitorStatus() { calls.push(['monitor-status']); return { status: 'degraded' }; },
    }, () => true);
    expect(ipc.handlers.get('material-events:list')?.({}, { query: '2330', unreadOnly: true, watchedOnly: true, eventIds: ['event-1'] })).toEqual([]);
    expect(ipc.handlers.get('material-events:detail')?.({}, { id: 'event-1' })).toEqual({ id: 'event-1' });
    expect(ipc.handlers.get('material-events:mark-read')?.({}, { id: 'event-1' })).toEqual({ id: 'event-1' });
    expect(ipc.handlers.get('material-events:soft-delete')?.({}, { id: 'event-1' })).toEqual({ id: 'event-1' });
    expect(ipc.handlers.get('material-events:status')?.({})).toEqual({ status: 'degraded' });
    expect(calls).toEqual([['list', { query: '2330', unreadOnly: true, watchedOnly: true, eventIds: ['event-1'] }], ['detail', 'event-1'], ['read', 'event-1'], ['delete', 'event-1'], ['monitor-status']]);
    dispose();
    expect(ipc.handlers.size).toBe(0);
  });

  it('rejects an untrusted sender and invalid IDs or filters', () => {
    const ipc = fakeIpc();
    registerMaterialIpc(ipc, { list() { return []; }, detail(id) { return id; }, markRead(id) { return id; }, softDelete(id) { return id; }, monitorStatus() { return null; } }, () => false);
    expect(() => ipc.handlers.get('material-events:list')?.({}, {})).toThrow(/不受信任/);
    expect(() => ipc.handlers.get('material-events:soft-delete')?.({}, { id: 'event-1' })).toThrow(/不受信任/);
    registerMaterialIpc(ipc, { list() { return []; }, detail(id) { return id; }, markRead(id) { return id; }, softDelete(id) { return id; }, monitorStatus() { return null; } }, () => true);
    expect(() => ipc.handlers.get('material-events:detail')?.({}, { id: '../token.json' })).toThrow(/ID/);
    expect(() => ipc.handlers.get('material-events:soft-delete')?.({}, { id: '../token.json' })).toThrow(/ID/);
    expect(() => ipc.handlers.get('material-events:list')?.({}, { unreadOnly: 'yes' })).toThrow(/條件/);
    expect(() => ipc.handlers.get('material-events:list')?.({}, { watchedOnly: 'yes' })).toThrow(/條件/);
    expect(() => ipc.handlers.get('material-events:list')?.({}, { eventIds: ['../token.json'] })).toThrow(/條件/);
  });
});
