import type { LoginStartupSettings } from './login-startup';

interface IpcPort {
  handle(channel: string, listener: (event: unknown, payload?: unknown) => unknown): void;
  removeHandler?(channel: string): void;
}

export function registerLoginStartupIpc(
  ipc: IpcPort,
  controller: { getSettings(): LoginStartupSettings; saveSettings(value: LoginStartupSettings): LoginStartupSettings },
  isTrustedSender: (event: unknown) => boolean,
): () => void {
  const handlers = new Map<string, (payload?: unknown) => unknown>([
    ['desktop:get-login-startup', () => controller.getSettings()],
    ['desktop:save-login-startup', (payload) => {
      if (!payload || typeof payload !== 'object' || Array.isArray(payload)
        || !('enabled' in payload) || typeof payload.enabled !== 'boolean'
        || !('startHidden' in payload) || typeof payload.startHidden !== 'boolean') {
        throw new Error('登入啟動設定格式無效');
      }
      return controller.saveSettings(payload as LoginStartupSettings);
    }],
  ]);
  for (const [channel, handler] of handlers) {
    ipc.handle(channel, (event, payload) => {
      if (!isTrustedSender(event)) throw new Error('拒絕不受信任的 renderer IPC 呼叫');
      return handler(payload);
    });
  }
  return () => handlers.forEach((_handler, channel) => ipc.removeHandler?.(channel));
}
