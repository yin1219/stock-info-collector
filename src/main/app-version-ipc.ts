interface IpcPort {
  handle(channel: string, listener: (event: unknown) => unknown): void;
  removeHandler?(channel: string): void;
}

export function registerAppVersionIpc(
  ipc: IpcPort,
  app: { getVersion(): string },
  isTrustedSender: (event: unknown) => boolean,
): () => void {
  const channel = 'desktop:get-app-version';
  ipc.handle(channel, (event) => {
    if (!isTrustedSender(event)) throw new Error('拒絕不受信任的 renderer IPC 呼叫');
    return app.getVersion();
  });
  return () => ipc.removeHandler?.(channel);
}
