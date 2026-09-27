import { describe, expect, it } from 'vitest';
import { registerDisclosureIpc } from '../../src/main/disclosure-ipc';

function fakeIpc() {
  const handlers = new Map<string, (event: unknown, payload?: unknown) => unknown>();
  return {
    handlers,
    handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown) { handlers.set(channel, listener); },
    removeHandler(channel: string) { handlers.delete(channel); },
  };
}

describe('default-disclosure-monitoring / IPC', () => {
  it('serves locally stored all-market rows through a trusted query', async () => {
    const ipc = fakeIpc();
    const seen: unknown[] = [];
    const dispose = registerDisclosureIpc(ipc, { list(input) { seen.push(input); return []; }, monitorStatus() { return { status: 'degraded' }; } }, () => true);
    expect(ipc.handlers.get('disclosures:list')?.({ from: 'main-window' }, { disclosureDate: '2026-09-26', market: 'TPEX' })).toEqual([]);
    expect(ipc.handlers.get('disclosures:status')?.({ from: 'main-window' })).toEqual({ status: 'degraded' });
    expect(seen).toEqual([{ disclosureDate: '2026-09-26', market: 'TPEX' }]);
    dispose();
    expect(ipc.handlers.size).toBe(0);
  });

  it('rejects untrusted senders and malformed filters', async () => {
    const ipc = fakeIpc();
    registerDisclosureIpc(ipc, { list() { return []; }, monitorStatus() { return null; } }, () => false);
    expect(() => ipc.handlers.get('disclosures:list')?.({}, {})).toThrow(/不受信任/);
    registerDisclosureIpc(ipc, { list() { return []; }, monitorStatus() { return null; } }, () => true);
    expect(() => ipc.handlers.get('disclosures:list')?.({}, { disclosureDate: '../../other' })).toThrow(/日期/);
    expect(() => ipc.handlers.get('disclosures:list')?.({}, { market: 'NYSE' })).toThrow(/市場/);
  });
});
