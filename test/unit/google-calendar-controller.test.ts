import { describe, expect, it, vi } from 'vitest';
import { createGoogleCalendarController } from '../../src/main/google-calendar-controller';

function setup() {
  const session = {
    getStatus: vi.fn(async () => ({ status: 'connected' as const })),
    configure: vi.fn(async () => undefined),
    connect: vi.fn(async () => undefined),
    disconnect: vi.fn(async () => undefined),
    markReauthorizationRequired: vi.fn(async () => undefined),
    migrateLegacyToken: vi.fn(async () => undefined),
  };
  const conferences = [{ id: 'conf-1', companyId: 'company-1', stockCode: '2330', companyName: '台積電', sourceKey: 'mops:conf-1', startsAt: '2026-10-01T02:00:00.000Z', location: '台北', content: '法說會' }];
  const repositories = {
    conferences: { list: vi.fn(() => conferences) },
    calendarSyncs: { find: vi.fn(() => ({ status: 'pending' as const, lastError: null })) },
    watchlist: { list: vi.fn(() => [{ stockCode: '2330', active: true }]) },
  };
  const syncConferences = vi.fn(async () => [{ status: 'synced' }, { status: 'failed', errorMessage: 'Google Calendar 授權已失效，請重新授權', authorizationRequired: true }]);
  const fetchConferences = vi.fn(async () => [{ CompId: '2330' }]);
  const selectCredentials = vi.fn(async () => ({ installed: { client_id: 'client-id' } }));
  const confirmLegacyTokenMigration = vi.fn(async () => true);
  const selectLegacyToken = vi.fn(async () => ({ type: 'authorized_user' }));
  const controller = createGoogleCalendarController({ session, repositories, fetchConferences, syncConferences, selectCredentials, confirmLegacyTokenMigration, selectLegacyToken });
  return { controller, session, repositories, syncConferences, fetchConferences, selectCredentials, confirmLegacyTokenMigration, selectLegacyToken };
}

describe('conference-calendar-sync / main-process controller', () => {
  it('reports calendar sync state without exposing saved credentials', async () => {
    const { controller } = setup();
    await expect(controller.status()).resolves.toEqual({ status: 'connected' });
    expect(controller.list()).toEqual([expect.objectContaining({ id: 'conf-1', syncStatus: 'pending' })]);
  });

  it('fetches active-company conferences and syncs a batch in main with mock provider ports', async () => {
    const { controller, fetchConferences, syncConferences, session } = setup();
    await expect(controller.sync()).resolves.toMatchObject({ status: 'reauthorization-required', results: [{ status: 'synced' }, { status: 'failed' }] });
    expect(session.markReauthorizationRequired).toHaveBeenCalledOnce();
    expect(fetchConferences).toHaveBeenCalledWith(['2330']);
    expect(syncConferences).toHaveBeenCalledWith([{ CompId: '2330' }]);
  });

  it('does not fetch or write when Google requires authorization', async () => {
    const { controller, session, fetchConferences, syncConferences } = setup();
    session.getStatus.mockResolvedValue({ status: 'reauthorization-required' });

    await expect(controller.sync()).rejects.toThrow('Google Calendar 需要重新授權');
    expect(fetchConferences).not.toHaveBeenCalled();
    expect(syncConferences).not.toHaveBeenCalled();
  });

  it('forwards connect, reconnect and disconnect only through the main-process session', async () => {
    const { controller, session } = setup();
    await controller.connect();
    await controller.disconnect();
    expect(session.connect).toHaveBeenCalledOnce();
    expect(session.disconnect).toHaveBeenCalledOnce();
  });

  it('imports only a user-selected OAuth file and treats cancellation as a no-op', async () => {
    const { controller, session, selectCredentials } = setup();
    await expect(controller.importCredentials()).resolves.toBe(true);
    expect(session.configure).toHaveBeenCalledWith({ installed: { client_id: 'client-id' } });
    selectCredentials.mockResolvedValueOnce(null);
    await expect(controller.importCredentials()).resolves.toBe(false);
    expect(session.configure).toHaveBeenCalledOnce();
  });

  it('requires explicit confirmation and selection before reading a legacy token for migration', async () => {
    const { controller, session, confirmLegacyTokenMigration, selectLegacyToken } = setup();
    confirmLegacyTokenMigration.mockResolvedValueOnce(false);
    await expect(controller.migrateLegacyToken()).resolves.toBe(false);
    expect(selectLegacyToken).not.toHaveBeenCalled();
    confirmLegacyTokenMigration.mockResolvedValueOnce(true);
    selectLegacyToken.mockResolvedValueOnce(null);
    await expect(controller.migrateLegacyToken()).resolves.toBe(false);
    expect(session.migrateLegacyToken).not.toHaveBeenCalled();
    await expect(controller.migrateLegacyToken()).resolves.toBe(true);
    expect(session.migrateLegacyToken).toHaveBeenCalledWith({ type: 'authorized_user' });
  });
});
