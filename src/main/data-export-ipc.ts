interface IpcPort {
  handle(channel: string, listener: (event: unknown) => unknown): void;
  removeHandler?(channel: string): void;
}

export function registerDataExportIpc(
  ipc: IpcPort,
  controller: { exportUserData(): Promise<unknown> },
  isTrustedSender: (event: unknown) => boolean,
): () => void {
  ipc.handle('data:export', (event) => {
    if (!isTrustedSender(event)) throw new Error('拒絕不受信任的 renderer IPC 呼叫');
    return controller.exportUserData();
  });
  return () => ipc.removeHandler?.('data:export');
}
