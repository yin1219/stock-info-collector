export interface IpcInvoker {
  invoke(channel: string, payload?: unknown): Promise<unknown>;
  on?(channel: string, listener: (event: unknown, payload: unknown) => void): unknown;
  removeListener?(channel: string, listener: (event: unknown, payload: unknown) => void): unknown;
}

export type ReporterRoute =
  | { type: 'event-detail'; eventId: string }
  | { type: 'event-list'; eventIds: string[] }
  | { type: 'source-status' }
  | { type: 'disclosure-list' }
  | { type: 'test-notification' };

export interface ScheduleSettings {
  monitoring: { enabled: boolean; start: string; end: string; intervalMinutes: 15 | 30 | 60 | 120 };
  disclosure: { enabled: boolean; runAt: string };
  notifyEmptyDefaultDisclosures: boolean;
}

export interface LoginStartupSettings {
  enabled: boolean;
  startHidden: boolean;
  blockedByWindows?: boolean;
}

export type GoogleCalendarStatus = { status: 'not-configured' | 'disconnected' | 'connected' | 'reauthorization-required' };
export type DataExportResult = { status: 'cancelled' } | { status: 'exported'; destination: string };

function isReporterRoute(value: unknown): value is ReporterRoute {
  if (!value || typeof value !== 'object' || !('type' in value)) return false;
  const route = value as Record<string, unknown>;
  if (route.type === 'event-detail') return typeof route.eventId === 'string' && route.eventId.length > 0;
  if (route.type === 'event-list') return Array.isArray(route.eventIds) && route.eventIds.every((id) => typeof id === 'string' && id.length > 0);
  if (route.type === 'source-status') return true;
  return route.type === 'disclosure-list' || route.type === 'test-notification';
}

export interface ReporterApi {
  getVersion(): string;
  listWatchlist(filter?: { query?: string; activeOnly?: boolean }): Promise<unknown>;
  addWatchlist(input: { market: 'TWSE' | 'TPEX'; stockCode: string; category?: string; notes?: string }): Promise<unknown>;
  updateWatchlist(input: { companyId: string; category: string; notes: string }): Promise<unknown>;
  setWatchlistActive(input: { companyId: string; active: boolean }): Promise<unknown>;
  removeWatchlist(input: { companyId: string }): Promise<unknown>;
  previewWatchlistImport(configText: string): Promise<unknown>;
  applyWatchlistImport(configText: string): Promise<unknown>;
  listDisclosures(filter?: { disclosureDate?: string; market?: 'TWSE' | 'TPEX' }): Promise<unknown>;
  getDisclosureMonitorStatus(): Promise<unknown>;
  listMaterialEvents(filter?: { query?: string; unreadOnly?: boolean; watchedOnly?: boolean; eventIds?: string[] }): Promise<unknown>;
  getMaterialEvent(id: string): Promise<unknown>;
  markMaterialEventRead(id: string): Promise<unknown>;
  softDeleteMaterialEvent(id: string): Promise<unknown>;
  getMaterialMonitorStatus(): Promise<unknown>;
  getScheduleSettings(): Promise<ScheduleSettings>;
  saveScheduleSettings(settings: ScheduleSettings): Promise<ScheduleSettings>;
  getScheduleStatus(): Promise<unknown>;
  runScheduledCheck(kind: 'material' | 'disclosure'): Promise<unknown>;
  sendTestNotification(): Promise<{ status: 'requested' | 'suppressed' }>;
  scheduleTestNotification(): Promise<{ status: 'scheduled'; scheduledAt: string } | { status: 'suppressed' }>;
  exportUserData(): Promise<DataExportResult>;
  getLoginStartupSettings(): Promise<LoginStartupSettings>;
  saveLoginStartupSettings(settings: LoginStartupSettings): Promise<LoginStartupSettings>;
  getGoogleCalendarStatus(): Promise<GoogleCalendarStatus>;
  importGoogleCredentials(): Promise<boolean>;
  migrateLegacyGoogleToken(): Promise<boolean>;
  connectGoogleCalendar(): Promise<GoogleCalendarStatus>;
  disconnectGoogleCalendar(): Promise<GoogleCalendarStatus>;
  listConferences(): Promise<unknown[]>;
  syncConferences(): Promise<unknown>;
  onNotificationRoute(listener: (route: ReporterRoute) => void): () => void;
}

export function createReporterApi(version: string, ipc: IpcInvoker): ReporterApi {
  const api: ReporterApi = {
    getVersion: () => version,
    listWatchlist: (filter = {}) => ipc.invoke('watchlist:list', filter),
    addWatchlist: (input) => ipc.invoke('watchlist:add', input),
    updateWatchlist: (input) => ipc.invoke('watchlist:update', input),
    setWatchlistActive: (input) => ipc.invoke('watchlist:set-active', input),
    removeWatchlist: (input) => ipc.invoke('watchlist:remove', input),
    previewWatchlistImport: (configText) => ipc.invoke('watchlist:import-preview', { configText }),
    applyWatchlistImport: (configText) => ipc.invoke('watchlist:import-apply', { configText }),
    listDisclosures: (filter = {}) => ipc.invoke('disclosures:list', filter),
    getDisclosureMonitorStatus: () => ipc.invoke('disclosures:status'),
    listMaterialEvents: (filter = {}) => ipc.invoke('material-events:list', filter),
    getMaterialEvent: (id) => ipc.invoke('material-events:detail', { id }),
    markMaterialEventRead: (id) => ipc.invoke('material-events:mark-read', { id }),
    softDeleteMaterialEvent: (id) => ipc.invoke('material-events:soft-delete', { id }),
    getMaterialMonitorStatus: () => ipc.invoke('material-events:status'),
    getScheduleSettings: () => ipc.invoke('schedule:get-settings') as Promise<ScheduleSettings>,
    saveScheduleSettings: (settings) => ipc.invoke('schedule:save-settings', settings) as Promise<ScheduleSettings>,
    getScheduleStatus: () => ipc.invoke('schedule:get-status'),
    runScheduledCheck: (kind) => ipc.invoke('schedule:run-now', { kind }),
    sendTestNotification: () => ipc.invoke('notification:send-test') as Promise<{ status: 'requested' | 'suppressed' }>,
    scheduleTestNotification: () => ipc.invoke('notification:schedule-test') as Promise<{ status: 'scheduled'; scheduledAt: string } | { status: 'suppressed' }>,
    exportUserData: () => ipc.invoke('data:export') as Promise<DataExportResult>,
    getLoginStartupSettings: () => ipc.invoke('desktop:get-login-startup') as Promise<LoginStartupSettings>,
    saveLoginStartupSettings: (settings) => ipc.invoke('desktop:save-login-startup', settings) as Promise<LoginStartupSettings>,
    getGoogleCalendarStatus: () => ipc.invoke('calendar:get-status') as Promise<GoogleCalendarStatus>,
    importGoogleCredentials: () => ipc.invoke('calendar:import-credentials') as Promise<boolean>,
    migrateLegacyGoogleToken: () => ipc.invoke('calendar:migrate-legacy-token') as Promise<boolean>,
    connectGoogleCalendar: () => ipc.invoke('calendar:connect') as Promise<GoogleCalendarStatus>,
    disconnectGoogleCalendar: () => ipc.invoke('calendar:disconnect') as Promise<GoogleCalendarStatus>,
    listConferences: () => ipc.invoke('calendar:list') as Promise<unknown[]>,
    syncConferences: () => ipc.invoke('calendar:sync'),
    onNotificationRoute(listener) {
      const handler = (_event: unknown, payload: unknown) => {
        if (isReporterRoute(payload)) listener(payload);
      };
      ipc.on?.('notification:navigate', handler);
      return () => { ipc.removeListener?.('notification:navigate', handler); };
    },
  };
  return Object.freeze(api);
}
