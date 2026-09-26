export interface MaterialIpcPort {
  handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown): void;
  removeHandler?(channel: string): void;
}

export interface MaterialReadService {
  list(filter: { query?: string; unreadOnly?: boolean; eventIds?: string[] }): unknown;
  detail(id: string): unknown;
  markRead(id: string): unknown;
}

export function registerMaterialIpc(
  ipcMain: MaterialIpcPort,
  service: MaterialReadService,
  isTrustedSender: (event: unknown) => boolean,
): () => void {
  const routes: Array<[string, (payload: unknown) => unknown]> = [
    ['material-events:list', (payload) => {
      if (payload !== undefined && (!payload || typeof payload !== 'object' || Array.isArray(payload))) throw new Error('重大訊息查詢條件格式無效');
      const value = (payload ?? {}) as Record<string, unknown>;
      if (value.query !== undefined && typeof value.query !== 'string') throw new Error('重大訊息查詢條件格式無效');
      if (value.unreadOnly !== undefined && typeof value.unreadOnly !== 'boolean') throw new Error('重大訊息查詢條件格式無效');
      if (value.eventIds !== undefined && (!Array.isArray(value.eventIds) || value.eventIds.length === 0 || value.eventIds.length > 100
        || !value.eventIds.every((id) => typeof id === 'string' && /^[A-Za-z0-9-]{1,80}$/.test(id)))) {
        throw new Error('重大訊息查詢條件格式無效');
      }
      return service.list({
        ...(typeof value.query === 'string' && value.query.trim() ? { query: value.query.trim().slice(0, 120) } : {}),
        ...(typeof value.unreadOnly === 'boolean' ? { unreadOnly: value.unreadOnly } : {}),
        ...(Array.isArray(value.eventIds) ? { eventIds: value.eventIds as string[] } : {}),
      });
    }],
    ['material-events:detail', (payload) => service.detail(readId(payload))],
    ['material-events:mark-read', (payload) => service.markRead(readId(payload))],
  ];
  for (const [channel, handler] of routes) {
    ipcMain.handle(channel, (event, payload) => {
      if (!isTrustedSender(event)) throw new Error('拒絕不受信任的 renderer IPC 呼叫');
      return handler(payload);
    });
  }
  return () => routes.forEach(([channel]) => ipcMain.removeHandler?.(channel));
}

function readId(payload: unknown): string {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload) || !('id' in payload)
    || typeof payload.id !== 'string' || !/^[A-Za-z0-9-]{1,80}$/.test(payload.id)) {
    throw new Error('重大訊息 ID 格式無效');
  }
  return payload.id;
}
