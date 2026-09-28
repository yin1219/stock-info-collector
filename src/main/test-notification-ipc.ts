interface IpcPort {
  handle(channel: string, listener: (event: unknown) => unknown): void;
  removeHandler?(channel: string): void;
}

export function registerTestNotificationIpc(
  ipc: IpcPort,
  controller: {
    send(): Promise<{ status: 'requested' | 'suppressed' }>;
    schedule(): { status: 'scheduled'; scheduledAt: string } | { status: 'suppressed' };
  },
  isTrustedSender: (event: unknown) => boolean,
  onFailure: (error: unknown) => void,
): () => void {
  ipc.handle('notification:send-test', (event) => {
    if (!isTrustedSender(event)) throw new Error('拒絕不受信任的 renderer IPC 呼叫');
    return controller.send().catch((error: unknown) => {
      onFailure(error);
      throw error;
    });
  });
  ipc.handle('notification:schedule-test', (event) => {
    if (!isTrustedSender(event)) throw new Error('拒絕不受信任的 renderer IPC 呼叫');
    return Promise.resolve(controller.schedule());
  });
  return () => {
    ipc.removeHandler?.('notification:send-test');
    ipc.removeHandler?.('notification:schedule-test');
  };
}
