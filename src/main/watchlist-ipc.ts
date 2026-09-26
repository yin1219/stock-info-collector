export interface IpcMainPort {
  handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown): void;
  removeHandler?(channel: string): void;
}

export interface WatchlistIpcService {
  list(filter: { query?: string; activeOnly?: boolean }): unknown;
  add(input: { market: 'TWSE' | 'TPEX'; stockCode: string; category?: string; notes?: string }): unknown;
  update(input: { companyId: string; category: string; notes: string }): unknown;
  setActive(input: { companyId: string; active: boolean }): unknown;
  remove(input: { companyId: string }): unknown;
  previewImport(input: { configText: string }): unknown;
  applyImport(input: { configText: string }): unknown;
}

function asRecord(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`${label}資料格式無效`);
  return value as Record<string, unknown>;
}

function companyId(value: unknown): string {
  const record = asRecord(value, '關注公司');
  if (typeof record.companyId !== 'string' || !record.companyId.trim()) throw new Error('缺少有效的 companyId');
  return record.companyId;
}

export function registerWatchlistIpc(
  ipcMain: IpcMainPort,
  service: WatchlistIpcService,
  isTrustedSender: (event: unknown) => boolean,
): () => void {
  const routes: Array<[string, (payload: unknown) => unknown]> = [
    ['watchlist:list', (payload) => {
      const record = payload == null ? {} : asRecord(payload, '查詢條件');
      return service.list({
        ...(typeof record.query === 'string' && record.query.trim() ? { query: record.query.trim().slice(0, 120) } : {}),
        ...(typeof record.activeOnly === 'boolean' ? { activeOnly: record.activeOnly } : {}),
      });
    }],
    ['watchlist:add', (payload) => {
      const record = asRecord(payload, '新增關注公司');
      if (record.market !== 'TWSE' && record.market !== 'TPEX') throw new Error('市場必須是 TWSE 或 TPEX');
      if (typeof record.stockCode !== 'string') throw new Error('股票代號格式無效');
      if (record.category !== undefined && typeof record.category !== 'string') throw new Error('分類格式無效');
      if (record.notes !== undefined && typeof record.notes !== 'string') throw new Error('備註格式無效');
      return service.add({
        market: record.market,
        stockCode: record.stockCode,
        ...(typeof record.category === 'string' ? { category: record.category } : {}),
        ...(typeof record.notes === 'string' ? { notes: record.notes } : {}),
      });
    }],
    ['watchlist:update', (payload) => {
      const record = asRecord(payload, '更新關注公司');
      if (typeof record.category !== 'string' || typeof record.notes !== 'string') throw new Error('分類與備註格式無效');
      return service.update({ companyId: companyId(payload), category: record.category, notes: record.notes });
    }],
    ['watchlist:set-active', (payload) => {
      const record = asRecord(payload, '關注狀態');
      if (typeof record.active !== 'boolean') throw new Error('active 必須是布林值');
      return service.setActive({ companyId: companyId(payload), active: record.active });
    }],
    ['watchlist:remove', (payload) => service.remove({ companyId: companyId(payload) })],
    ['watchlist:import-preview', (payload) => {
      const record = asRecord(payload, '匯入設定');
      if (typeof record.configText !== 'string') throw new Error('匯入設定內容無效');
      return service.previewImport({ configText: record.configText });
    }],
    ['watchlist:import-apply', (payload) => {
      const record = asRecord(payload, '匯入設定');
      if (typeof record.configText !== 'string') throw new Error('匯入設定內容無效');
      return service.applyImport({ configText: record.configText });
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
