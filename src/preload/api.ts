export interface IpcInvoker {
  invoke(channel: string, payload?: unknown): Promise<unknown>;
  on?(channel: string, listener: (event: unknown, payload: unknown) => void): unknown;
  removeListener?(channel: string, listener: (event: unknown, payload: unknown) => void): unknown;
}

export type ReporterRoute =
  | { type: 'event-detail'; eventId: string }
  | { type: 'event-list'; eventIds: string[] }
  | { type: 'source-status' }
  | { type: 'market-records' }
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
export type ConferenceTaskStatus = { kind: 'legacy' | 'legacy-disabled' | 'v2' | 'blocked' | 'create-v2' | 'upgrade-v2-time'; reason?: string; configPath?: string; legacyConfigPath?: string; setupReady?: boolean; lastTaskResult?: number };
export type LegacyConferenceConfigReview = { configPath: string; hash: string; currentCodes: string[]; activeCodes: string[]; confirmed: boolean };
export type LegacyConferenceConfigApply = { status: 'applied' | 'unchanged'; stockNumbers: string[]; backupPath?: string };
export type DataExportResult = { status: 'cancelled' } | { status: 'exported'; destination: string };

function isReporterRoute(value: unknown): value is ReporterRoute {
  if (!value || typeof value !== 'object' || !('type' in value)) return false;
  const route = value as Record<string, unknown>;
  if (route.type === 'event-detail') return typeof route.eventId === 'string' && route.eventId.length > 0;
  if (route.type === 'event-list') return Array.isArray(route.eventIds) && route.eventIds.every((id) => typeof id === 'string' && id.length > 0);
  if (route.type === 'source-status') return true;
  if (route.type === 'market-records') return true;
  return route.type === 'disclosure-list' || route.type === 'test-notification';
}

export interface ReporterApi {
  getMarketRecords(): Promise<unknown>;
  refreshMarketRecords(): Promise<unknown>;
  importMarketRecords(kind: 'foreign' | 'trust' | 'dealer' | 'index' | 'csv'): Promise<unknown>;
  getVersion(): Promise<string>;
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
  getConferenceTaskStatus(): Promise<ConferenceTaskStatus>;
  ensureConferenceTask(): Promise<ConferenceTaskStatus>;
  switchFromDisabledLegacy(): Promise<ConferenceTaskStatus>;
  reviewLegacyConferenceConfig(): Promise<LegacyConferenceConfigReview>;
  confirmLegacyConferenceConfig(hash: string): Promise<LegacyConferenceConfigApply>;
  onNotificationRoute(listener: (route: ReporterRoute) => void): () => void;
}

export function createReporterApi(ipc: IpcInvoker): ReporterApi {
  const api: ReporterApi = {
    getMarketRecords: () => ipc.invoke('market-records:get'),
    refreshMarketRecords: () => ipc.invoke('market-records:refresh'),
    importMarketRecords: (kind) => ipc.invoke('market-records:import', kind),
    getVersion: () => ipc.invoke('desktop:get-app-version') as Promise<string>,
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
    getConferenceTaskStatus: () => ipc.invoke('conference-task:status') as Promise<ConferenceTaskStatus>,
    ensureConferenceTask: () => ipc.invoke('conference-task:ensure') as Promise<ConferenceTaskStatus>,
    switchFromDisabledLegacy: () => ipc.invoke('conference-task:switch-disabled-legacy', { confirmation: 'SWITCH_TO_V2_DAILY' }) as Promise<ConferenceTaskStatus>,
    reviewLegacyConferenceConfig: () => ipc.invoke('conference-task:review-legacy') as Promise<LegacyConferenceConfigReview>,
    confirmLegacyConferenceConfig: (hash) => ipc.invoke('conference-task:confirm-legacy', { hash }) as Promise<LegacyConferenceConfigApply>,
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
