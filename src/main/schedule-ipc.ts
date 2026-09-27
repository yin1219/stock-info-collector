import type { ScheduleSettings } from '../services/schedule-settings';

export interface ScheduleIpcPort {
  handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown): void;
  removeHandler?(channel: string): void;
}

export interface ScheduleService {
  getSettings(): ScheduleSettings;
  saveSettings(value: ScheduleSettings): ScheduleSettings;
  getStatus(): unknown;
  runNow(kind: 'material' | 'disclosure'): unknown;
}

export function registerScheduleIpc(ipcMain: ScheduleIpcPort, service: ScheduleService, isTrustedSender: (event: unknown) => boolean): () => void {
  const routes: Array<[string, (payload: unknown) => unknown]> = [
    ['schedule:get-settings', () => service.getSettings()],
    ['schedule:get-status', () => service.getStatus()],
    ['schedule:save-settings', (payload) => service.saveSettings(payload as ScheduleSettings)],
    ['schedule:run-now', (payload) => {
      if (!payload || typeof payload !== 'object' || Array.isArray(payload)
        || !('kind' in payload) || (payload.kind !== 'material' && payload.kind !== 'disclosure')) {
        throw new Error('立即檢查工作種類無效');
      }
      return service.runNow(payload.kind);
    }],
  ];
  for (const [channel, handler] of routes) {
    ipcMain.handle(channel, (event, payload) => {
      if (!isTrustedSender(event)) throw new Error('拒絕不受信任的 renderer IPC 呼叫');
      return handler(payload);
    });
  }
  return () => routes.forEach(([channel]) => ipcMain.removeHandler?.(channel));
}
