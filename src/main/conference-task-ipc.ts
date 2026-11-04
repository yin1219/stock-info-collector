import type { IpcMainPort } from './watchlist-ipc';

export interface ConferenceTaskIpcService {
  status(): unknown;
  ensure(): unknown;
  switchFromDisabledLegacy(): unknown;
  reviewLegacy(): unknown;
  confirmLegacy(hash: string): unknown;
}

export function registerConferenceTaskIpc(
  ipcMain: IpcMainPort,
  service: ConferenceTaskIpcService,
  isTrustedSender: (event: unknown) => boolean,
): () => void {
  const routes: Array<[string, (payload: unknown) => unknown]> = [
    ['conference-task:status', () => service.status()],
    ['conference-task:ensure', () => service.ensure()],
    ['conference-task:switch-disabled-legacy', (payload) => {
      if (!payload || typeof payload !== 'object' || Array.isArray(payload)
        || !('confirmation' in payload) || payload.confirmation !== 'SWITCH_TO_V2_DAILY') {
        throw new Error('切換 v2 每日工作需要明確確認');
      }
      return service.switchFromDisabledLegacy();
    }],
    ['conference-task:review-legacy', () => service.reviewLegacy()],
    ['conference-task:confirm-legacy', (payload) => {
      if (!payload || typeof payload !== 'object' || Array.isArray(payload)
        || !('hash' in payload) || typeof payload.hash !== 'string'
        || !/^[a-f0-9]{64}$/i.test(payload.hash)) {
        throw new Error('缺少有效的 v1 設定核對版本');
      }
      return service.confirmLegacy(payload.hash);
    }],
  ];
  for (const [channel, handler] of routes) {
    ipcMain.handle(channel, async (event, payload) => {
      if (!isTrustedSender(event)) throw new Error('拒絕不受信任的 renderer IPC 呼叫');
      return handler(payload);
    });
  }
  return () => routes.forEach(([channel]) => ipcMain.removeHandler?.(channel));
}
