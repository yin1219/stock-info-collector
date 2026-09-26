export interface DisclosureIpcPort {
  handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown): void;
  removeHandler?(channel: string): void;
}

export interface DisclosureReadService {
  list(filter: { disclosureDate?: string; market?: 'TWSE' | 'TPEX' }): unknown;
}

export function registerDisclosureIpc(
  ipcMain: DisclosureIpcPort,
  service: DisclosureReadService,
  isTrustedSender: (event: unknown) => boolean,
): () => void {
  const channel = 'disclosures:list';
  ipcMain.handle(channel, (event, payload) => {
    if (!isTrustedSender(event)) throw new Error('拒絕不受信任的 renderer IPC 呼叫');
    if (payload !== undefined && (!payload || typeof payload !== 'object' || Array.isArray(payload))) {
      throw new Error('違約交割查詢條件格式無效');
    }
    const value = (payload ?? {}) as Record<string, unknown>;
    if (value.disclosureDate !== undefined && (typeof value.disclosureDate !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value.disclosureDate))) {
      throw new Error('違約交割查詢日期格式無效');
    }
    if (value.market !== undefined && value.market !== 'TWSE' && value.market !== 'TPEX') {
      throw new Error('違約交割市場只允許 TWSE 或 TPEX');
    }
    return service.list({
      ...(typeof value.disclosureDate === 'string' ? { disclosureDate: value.disclosureDate } : {}),
      ...(value.market === 'TWSE' || value.market === 'TPEX' ? { market: value.market } : {}),
    });
  });
  return () => ipcMain.removeHandler?.(channel);
}
