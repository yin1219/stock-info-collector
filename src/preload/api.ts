export interface IpcInvoker {
  invoke(channel: string, payload?: unknown): Promise<unknown>;
  on?(channel: string, listener: (event: unknown, payload: unknown) => void): unknown;
  removeListener?(channel: string, listener: (event: unknown, payload: unknown) => void): unknown;
}

export type ReporterRoute =
  | { type: 'event-detail'; eventId: string }
  | { type: 'event-list'; eventIds: string[] }
  | { type: 'disclosure-list' };

function isReporterRoute(value: unknown): value is ReporterRoute {
  if (!value || typeof value !== 'object' || !('type' in value)) return false;
  const route = value as Record<string, unknown>;
  if (route.type === 'event-detail') return typeof route.eventId === 'string' && route.eventId.length > 0;
  if (route.type === 'event-list') return Array.isArray(route.eventIds) && route.eventIds.every((id) => typeof id === 'string' && id.length > 0);
  return route.type === 'disclosure-list';
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
  listMaterialEvents(filter?: { query?: string; unreadOnly?: boolean; eventIds?: string[] }): Promise<unknown>;
  getMaterialEvent(id: string): Promise<unknown>;
  markMaterialEventRead(id: string): Promise<unknown>;
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
    listMaterialEvents: (filter = {}) => ipc.invoke('material-events:list', filter),
    getMaterialEvent: (id) => ipc.invoke('material-events:detail', { id }),
    markMaterialEventRead: (id) => ipc.invoke('material-events:mark-read', { id }),
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
