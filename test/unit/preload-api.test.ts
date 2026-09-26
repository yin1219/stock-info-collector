import { describe, expect, it } from 'vitest';
import { createReporterApi } from '../../src/preload/api';

describe('security architecture / preload exposes narrow typed watchlist APIs', () => {
  it('maps UI operations to fixed IPC channels without exposing invoke or filesystem primitives', async () => {
    const calls: Array<{ channel: string; payload?: unknown }> = [];
    const api = createReporterApi('44.0.0', {
      invoke(channel, payload) { calls.push({ channel, payload }); return Promise.resolve(channel); },
    });

    expect(api.getVersion()).toBe('44.0.0');
    await api.listWatchlist({ query: '2330', activeOnly: true });
    await api.addWatchlist({ market: 'TWSE', stockCode: '2330' });
    await api.updateWatchlist({ companyId: 'co-1', category: '半導體', notes: '追蹤' });
    await api.setWatchlistActive({ companyId: 'co-1', active: false });
    await api.removeWatchlist({ companyId: 'co-1' });
    await api.previewWatchlistImport('{"StockNumbers":"2330"}');
    await api.applyWatchlistImport('{"StockNumbers":"2330"}');

    expect(calls.map(({ channel }) => channel)).toEqual([
      'watchlist:list', 'watchlist:add', 'watchlist:update', 'watchlist:set-active',
      'watchlist:remove', 'watchlist:import-preview', 'watchlist:import-apply',
    ]);
    expect('invoke' in api).toBe(false);
    expect('readFile' in api).toBe(false);
  });

  it('accepts only typed internal notification routes and provides an unsubscribe handle', () => {
    let listener: ((_event: unknown, payload: unknown) => void) | undefined;
    let removed = false;
    const api = createReporterApi('44.0.0', {
      invoke() { return Promise.resolve(); },
      on(channel, next) { expect(channel).toBe('notification:navigate'); listener = next; return this; },
      removeListener(_channel, target) { removed = target === listener; return this; },
    });
    const routes: unknown[] = [];

    const unsubscribe = api.onNotificationRoute((route) => routes.push(route));
    listener?.({}, { type: 'event-detail', eventId: 'event-1' });
    listener?.({}, { type: 'event-detail', eventId: 123 });
    unsubscribe();

    expect(routes).toEqual([{ type: 'event-detail', eventId: 'event-1' }]);
    expect(removed).toBe(true);
  });
});
