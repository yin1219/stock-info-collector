import { describe, expect, it } from 'vitest';
import { createReporterApi } from '../../src/preload/api';

describe('security architecture / preload exposes narrow typed watchlist APIs', () => {
  it('maps UI operations to fixed IPC channels without exposing invoke or filesystem primitives', async () => {
    const calls: Array<{ channel: string; payload?: unknown }> = [];
    const api = createReporterApi({
      invoke(channel, payload) { calls.push({ channel, payload }); return Promise.resolve(channel); },
    });

    expect(await api.getVersion()).toBe('desktop:get-app-version');
    await api.listWatchlist({ query: '2330', activeOnly: true });
    await api.addWatchlist({ market: 'TWSE', stockCode: '2330' });
    await api.updateWatchlist({ companyId: 'co-1', category: '半導體', notes: '追蹤' });
    await api.setWatchlistActive({ companyId: 'co-1', active: false });
    await api.removeWatchlist({ companyId: 'co-1' });
    await api.previewWatchlistImport('{"StockNumbers":"2330"}');
    await api.applyWatchlistImport('{"StockNumbers":"2330"}');
    await api.getGoogleCalendarStatus();
    await api.importGoogleCredentials();
    await api.migrateLegacyGoogleToken();
    await api.connectGoogleCalendar();
    await api.listConferences();
    await api.syncConferences();
    await api.disconnectGoogleCalendar();
    await api.exportUserData();
    await api.sendTestNotification();
    await api.scheduleTestNotification();
    await api.softDeleteMaterialEvent('event-1');
    await api.getConferenceTaskStatus();
    await api.ensureConferenceTask();
    await api.switchFromDisabledLegacy();
    await api.reviewLegacyConferenceConfig();
    await api.confirmLegacyConferenceConfig('a'.repeat(64));

    expect(calls.map(({ channel }) => channel)).toEqual([
      'desktop:get-app-version', 'watchlist:list', 'watchlist:add', 'watchlist:update', 'watchlist:set-active',
      'watchlist:remove', 'watchlist:import-preview', 'watchlist:import-apply',
      'calendar:get-status', 'calendar:import-credentials', 'calendar:migrate-legacy-token', 'calendar:connect', 'calendar:list', 'calendar:sync', 'calendar:disconnect',
      'data:export', 'notification:send-test', 'notification:schedule-test', 'material-events:soft-delete',
      'conference-task:status', 'conference-task:ensure', 'conference-task:switch-disabled-legacy', 'conference-task:review-legacy', 'conference-task:confirm-legacy',
    ]);
    expect(calls.at(-1)?.payload).toEqual({ hash: 'a'.repeat(64) });
    expect('invoke' in api).toBe(false);
    expect('readFile' in api).toBe(false);
  });

  it('accepts only typed internal notification routes and provides an unsubscribe handle', () => {
    let listener: ((_event: unknown, payload: unknown) => void) | undefined;
    let removed = false;
    const api = createReporterApi({
      invoke() { return Promise.resolve(); },
      on(channel, next) { expect(channel).toBe('notification:navigate'); listener = next; return this; },
      removeListener(_channel, target) { removed = target === listener; return this; },
    });
    const routes: unknown[] = [];

    const unsubscribe = api.onNotificationRoute((route) => routes.push(route));
    listener?.({}, { type: 'event-detail', eventId: 'event-1' });
    listener?.({}, { type: 'event-detail', eventId: 123 });
    listener?.({}, { type: 'test-notification' });
    unsubscribe();

    expect(routes).toEqual([{ type: 'event-detail', eventId: 'event-1' }, { type: 'test-notification' }]);
    expect(removed).toBe(true);
  });
});
