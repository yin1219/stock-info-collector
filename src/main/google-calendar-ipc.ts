export interface GoogleCalendarIpcPort {
  handle(channel: string, listener: (event: unknown) => unknown): void;
  removeHandler?(channel: string): void;
}

export interface GoogleCalendarControllerPort {
  status(): unknown;
  list(): unknown;
  importCredentials(): unknown;
  migrateLegacyToken(): unknown;
  connect(): unknown;
  disconnect(): unknown;
  sync(): unknown;
}

export function registerGoogleCalendarIpc(
  ipcMain: GoogleCalendarIpcPort,
  controller: GoogleCalendarControllerPort,
  isTrustedSender: (event: unknown) => boolean,
  onFailure?: (channel: string, error: unknown) => void,
): () => void {
  const routes: Array<[string, () => unknown]> = [
    ['calendar:get-status', () => controller.status()],
    ['calendar:list', () => controller.list()],
    ['calendar:import-credentials', () => controller.importCredentials()],
    ['calendar:migrate-legacy-token', () => controller.migrateLegacyToken()],
    ['calendar:connect', () => controller.connect()],
    ['calendar:disconnect', () => controller.disconnect()],
    ['calendar:sync', () => controller.sync()],
  ];
  for (const [channel, handler] of routes) {
    ipcMain.handle(channel, (event) => {
      if (!isTrustedSender(event)) throw new Error('拒絕不受信任的 renderer IPC 呼叫');
      try {
        return Promise.resolve(handler()).catch((error: unknown) => {
          onFailure?.(channel, error);
          throw error;
        });
      } catch (error) {
        onFailure?.(channel, error);
        throw error;
      }
    });
  }
  return () => routes.forEach(([channel]) => ipcMain.removeHandler?.(channel));
}
