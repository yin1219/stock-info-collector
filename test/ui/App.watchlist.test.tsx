import { cleanup, fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import axe from 'axe-core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { ReporterApi } from '../../src/preload/api';
import { App } from '../../src/renderer/App';

const company = {
  id: 'watch-1', companyId: 'company-1', market: 'TWSE' as const, stockCode: '2330',
  name: '台積電', active: true, category: '半導體', notes: '追蹤法說會',
  createdAt: '2026-01-01T00:00:00.000Z', updatedAt: '2026-01-01T00:00:00.000Z',
};

function setupApi() {
  let rows = [company];
  const materialEvents = [{
    id: 'event-2', companyId: 'company-1', market: 'TWSE' as const, stockCode: '2330', companyName: '台積電',
    title: '營收更正公告', content: '更正後完整公告內容', source: 'mops', sourceUrl: 'https://mops.example/event-2',
    publishedAt: '2026-09-26T09:30:00.000Z', discoveredAt: '2026-09-26T09:35:00.000Z', eventType: 'correction',
    revisionOf: 'event-1', readAt: null,
  }, {
    id: 'event-1', companyId: 'company-1', market: 'TWSE' as const, stockCode: '2330', companyName: '台積電',
    title: '營運報告', content: '原公告完整內容', source: 'mops', sourceUrl: 'https://mops.example/event-1',
    publishedAt: '2026-09-25T09:30:00.000Z', discoveredAt: '2026-09-25T09:35:00.000Z', eventType: 'announcement',
    revisionOf: null, readAt: '2026-09-25T10:00:00.000Z',
  }, {
    id: 'event-3', companyId: 'company-3', market: 'TPEX' as const, stockCode: '6488', companyName: '其他公司',
    title: '其他重大公告', content: '未關注公司公告內容', source: 'mops', sourceUrl: 'https://mops.example/event-3',
    publishedAt: '2026-09-26T08:30:00.000Z', discoveredAt: '2026-09-26T08:35:00.000Z', eventType: 'announcement',
    revisionOf: null, readAt: null,
  }];
  let notificationRouteListener: ((route: { type: 'event-detail'; eventId: string } | { type: 'event-list'; eventIds: string[] } | { type: 'source-status' } | { type: 'disclosure-list' }) => void) | undefined;
  let scheduleSettings = {
    monitoring: { enabled: true, start: '00:00', end: '00:00', intervalMinutes: 60 as const },
    disclosure: { enabled: true, runAt: '18:30' },
    notifyEmptyDefaultDisclosures: true,
  };
  let loginStartupSettings = { enabled: false, startHidden: false };
  let googleCalendarStatus = { status: 'disconnected' as const };
  const conferences = [
    { id: 'conf-pending', companyId: 'company-1', stockCode: '2330', companyName: '台積電', startsAt: '2026-10-01T02:00:00.000Z', location: '台北', content: '法說會', sourceUrl: 'https://mops.twse.com.tw/mops/web/t100sb07_1?co_id=2330', syncStatus: 'pending', lastError: null },
    { id: 'conf-existing', companyId: 'company-1', stockCode: '2330', companyName: '台積電', startsAt: '2026-10-02T02:00:00.000Z', location: '線上', content: '法說會', syncStatus: 'existing', lastError: null },
    { id: 'conf-synced', companyId: 'company-1', stockCode: '2330', companyName: '台積電', startsAt: '2026-10-03T02:00:00.000Z', location: '線上', content: '法說會', syncStatus: 'synced', lastError: null },
    { id: 'conf-failed', companyId: 'company-1', stockCode: '2330', companyName: '台積電', startsAt: '2026-10-04T02:00:00.000Z', location: '台北', content: '法說會', syncStatus: 'failed', lastError: 'Google Calendar 需要重新授權' },
  ];
  const scheduleStatus = {
    material: { enabled: true, status: 'failed', lastSuccessAt: '2026-09-26T02:00:00.000Z', lastResult: 'failed', errorMessage: 'MOPS 暫時無法連線', nextRunAt: '2026-09-26T03:00:00.000Z' },
    disclosure: { enabled: true, status: 'complete', lastSuccessAt: '2026-09-25T10:30:00.000Z', lastResult: 'complete', errorMessage: null, nextRunAt: '2026-09-26T10:30:00.000Z' },
  };
  const api: ReporterApi = {
    getVersion: () => 'test',
    getMaterialMonitorStatus: vi.fn(async () => ({
      status: 'degraded', finishedAt: '2026-09-26T11:00:00.000Z', errorMessage: 'MOPS 暫時無法連線，RSS 備援已使用。',
      sources: [{ source: 'MOPS', status: 'failed' }, { source: 'MOPS-RSS', status: 'complete' }, { source: 'TWSE-reconciliation', status: 'complete' }],
    })),
    getGoogleCalendarStatus: vi.fn(async () => googleCalendarStatus),
    importGoogleCredentials: vi.fn(async () => true),
    migrateLegacyGoogleToken: vi.fn(async () => true),
    connectGoogleCalendar: vi.fn(async () => { googleCalendarStatus = { status: 'connected' }; return googleCalendarStatus; }),
    disconnectGoogleCalendar: vi.fn(async () => { googleCalendarStatus = { status: 'disconnected' }; return googleCalendarStatus; }),
    listConferences: vi.fn(async () => conferences),
    syncConferences: vi.fn(async () => ({ status: 'complete', results: [] })),
    getScheduleSettings: vi.fn(async () => scheduleSettings),
    saveScheduleSettings: vi.fn(async (value) => { scheduleSettings = value; return value; }),
    getScheduleStatus: vi.fn(async () => scheduleStatus),
    runScheduledCheck: vi.fn(async () => ({ status: 'started' })),
    exportUserData: vi.fn(async () => ({ status: 'cancelled' as const })),
    getLoginStartupSettings: vi.fn(async () => loginStartupSettings),
    saveLoginStartupSettings: vi.fn(async (value) => { loginStartupSettings = value; return value; }),
    listWatchlist: vi.fn(async (filter = {}) => rows.filter((row) =>
      (!filter.activeOnly || row.active)
      && (!filter.query || `${row.stockCode} ${row.name}`.includes(filter.query)),
    )),
    addWatchlist: vi.fn(async (input) => {
      rows = [{ ...company, ...input, id: 'watch-2', companyId: 'company-2' }, ...rows];
      return rows[0];
    }),
    updateWatchlist: vi.fn(async (input) => { rows = rows.map((row) => row.companyId === input.companyId ? { ...row, category: input.category, notes: input.notes } : row); return rows[0]; }),
    setWatchlistActive: vi.fn(async ({ active }) => { rows = rows.map((row) => ({ ...row, active })); return rows[0]; }),
    removeWatchlist: vi.fn(async ({ companyId }) => { rows = rows.filter((row) => row.companyId !== companyId); }),
    previewWatchlistImport: vi.fn(async () => ({ summary: { added: 1, skipped: 1, failed: 1 }, items: [] })),
    applyWatchlistImport: vi.fn(async () => ({ summary: { added: 1, skipped: 1, failed: 1 }, items: [] })),
    listDisclosures: vi.fn(async () => [
      { id: 'd1', market: 'TWSE', stockCode: '2330', companyName: '關注半導體', disclosureDate: '2026-09-26', brokerCode: 'B001', isWatched: true, contentJson: '{"amount":50000}' },
      { id: 'd2', market: 'TPEX', stockCode: '6488', companyName: '一般公司', disclosureDate: '2026-09-26', brokerCode: 'B003', isWatched: false, contentJson: '{"amount":10000}' },
    ]),
    getDisclosureMonitorStatus: vi.fn(async () => ({
      status: 'degraded', finishedAt: '2026-09-26T11:00:00.000Z', errorMessage: 'TPEX: 來源逾時',
      markets: {
        TWSE: { status: 'stale', dataDate: '2026-09-25', recordCount: 3, retryAt: '2026-09-26T19:00:00+08:00' },
        TPEX: { status: 'failed', dataDate: null, recordCount: 0, retryAt: null, errorMessage: '來源逾時' },
      },
    })),
    listMaterialEvents: vi.fn(async (filter = {}) => materialEvents.filter((event) =>
      (!filter.unreadOnly || !event.readAt)
      && (!filter.watchedOnly || rows.some((row) => row.companyId === event.companyId && row.active))
      && (!filter.eventIds || filter.eventIds.includes(event.id))
      && (!filter.query || `${event.stockCode} ${event.companyName} ${event.title} ${event.content}`.includes(filter.query)),
    )),
    onNotificationRoute: vi.fn((listener) => { notificationRouteListener = listener; return () => { notificationRouteListener = undefined; }; }),
    getMaterialEvent: vi.fn(async (id) => {
      const event = materialEvents.find((item) => item.id === id)!;
      return { ...event, relatedRevisions: [], original: event.revisionOf ? materialEvents.find((item) => item.id === event.revisionOf) : null };
    }),
    markMaterialEventRead: vi.fn(async (id) => {
      const event = materialEvents.find((item) => item.id === id)!;
      event.readAt = '2026-09-26T10:00:00.000Z';
      return event;
    }),
  };
  Object.defineProperty(window, 'reporterApi', { configurable: true, value: api });
  Object.defineProperty(window, 'sendReporterNotificationRoute', { configurable: true, value: (route: Parameters<NonNullable<typeof notificationRouteListener>>[0]) => notificationRouteListener?.(route) });
  return api;
}

describe('watchlist-management / UI', () => {
  afterEach(() => cleanup());

  it('removes a company from the visible list instead of showing it as disabled', async () => {
    const api = setupApi();
    const user = userEvent.setup();
    render(<App />);
    await user.click(screen.getByRole('tab', { name: '關注公司' }));
    expect(await screen.findByText('台積電')).toBeVisible();
    await user.click(screen.getByRole('button', { name: '移除台積電' }));
    await waitFor(() => expect(screen.queryByText('台積電')).not.toBeInTheDocument());
    expect(screen.getByText('目前沒有符合條件的公司。')).toBeVisible();
    expect(api.removeWatchlist).toHaveBeenCalledWith({ companyId: 'company-1' });
  });
  beforeEach(() => vi.restoreAllMocks());

  it('desktop-app-lifecycle / overview dashboard / summarizes saved watchlist and current monitor data', async () => {
    const api = setupApi();
    render(<App />);
    expect(screen.getAllByRole('tab').map((tab) => tab.textContent?.trim())).toEqual([
      '總覽', '關注公司', '重大訊息', '違約交割', '法說會', '設定',
    ]);
    await userEvent.click(await screen.findByRole('tab', { name: '總覽' }));
    expect(await screen.findByRole('heading', { name: '今日監控總覽' })).toBeVisible();
    expect(await screen.findByText('關注公司 1 家')).toBeVisible();
    expect(await screen.findByRole('heading', { name: '剛收到的重大訊息' })).toBeVisible();
    expect(await screen.findByRole('heading', { name: '監控排程' })).toBeVisible();
    expect(await screen.findByRole('heading', { name: '近期法說會' })).toBeVisible();
    expect(await screen.findByRole('heading', { name: '資料來源' })).toBeVisible();
    expect(api.getMaterialMonitorStatus).toHaveBeenCalled();
    expect(api.listDisclosures).toHaveBeenCalled();
    expect(api.listConferences).toHaveBeenCalled();
    expect(api.getScheduleStatus).toHaveBeenCalled();
  });

  it('desktop-app-lifecycle / overview dashboard / distinguishes an unrun monitor from a still-loading status', async () => {
    const api = setupApi();
    Object.defineProperty(api, 'getMaterialMonitorStatus', { value: vi.fn(async () => null) });
    Object.defineProperty(api, 'getDisclosureMonitorStatus', { value: vi.fn(async () => null) });
    render(<App />);
    await userEvent.click(await screen.findByRole('tab', { name: '總覽' }));

    expect(await screen.findByText('來源：尚未檢查')).toBeVisible();
    expect(screen.getAllByText('尚未檢查')).toHaveLength(2);
    expect(screen.queryByText('載入中…')).not.toBeInTheDocument();
  });

  it('local-data-management / export UI / lets the user export and reports success or failure', async () => {
    const api = setupApi() as ReporterApi & {
      exportUserData: ReturnType<typeof vi.fn>;
    };
    api.exportUserData = vi.fn()
      .mockResolvedValueOnce({ status: 'exported', destination: 'C:/Documents/reporter.json' })
      .mockRejectedValueOnce(new Error('無法寫入所選位置'));
    render(<App />);
    await userEvent.click(await screen.findByRole('tab', { name: '設定' }));

    await userEvent.click(await screen.findByRole('button', { name: '匯出資料' }));
    expect(api.exportUserData).toHaveBeenCalledOnce();
    expect(await screen.findByRole('status')).toHaveTextContent('資料已匯出至 C:/Documents/reporter.json');

    await userEvent.click(screen.getByRole('button', { name: '匯出資料' }));
    expect(await screen.findByText('無法寫入所選位置')).toHaveAttribute('role', 'alert');
  });

  it('notification-delivery / empty disclosure preference UI / saves whether zero-result notices are enabled', async () => {
    const api = setupApi();
    render(<App />);
    await userEvent.click(await screen.findByRole('tab', { name: '設定' }));
    const preference = await screen.findByRole('checkbox', { name: '收到本日無違約揭露通知' });
    expect(preference).toBeChecked();
    await userEvent.click(preference);
    await userEvent.click(screen.getByRole('button', { name: '儲存排程' }));

    await waitFor(() => expect(api.saveScheduleSettings).toHaveBeenCalledWith(expect.objectContaining({
      notifyEmptyDefaultDisclosures: false,
    })));
  });

  it('desktop-app-lifecycle / accessibility / exposes labelled controls and landmarks on primary screens', async () => {
    setupApi();
    render(<App />);
    for (const screenName of ['總覽', '關注公司', '重大訊息', '違約交割', '法說會', '設定']) {
      await userEvent.click(screen.getByRole('tab', { name: screenName }));
      await waitFor(async () => {
        const result = await axe.run(document.body, { runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] } });
        expect(result.violations.map(({ id, help }) => `${id}: ${help}`)).toEqual([]);
      });
    }
  });

  it('desktop-app-lifecycle / keyboard navigation / moves through the primary navigation without a pointer', async () => {
    setupApi();
    const user = userEvent.setup();
    render(<App />);
    await user.tab();
    expect(screen.getByRole('tab', { name: '總覽' })).toHaveFocus();
    await user.tab();
    expect(screen.getByRole('tab', { name: '關注公司' })).toHaveFocus();
    await user.keyboard('{Enter}');
    expect(await screen.findByRole('heading', { name: '關注公司' })).toBeVisible();
  });

  it('conference-calendar-sync / UI / shows sync outcomes and reauthorization action for saved conferences', async () => {
    const api = setupApi();
    render(<App />);
    await userEvent.click(await screen.findByRole('tab', { name: '法說會' }));
    expect(await screen.findByRole('heading', { name: '法說會與 Google Calendar' })).toBeVisible();
    expect(await screen.findByText('待同步')).toBeVisible();
    expect(screen.getByRole('link', { name: '開啟法說會來源' })).toHaveAttribute('href', 'https://mops.twse.com.tw/mops/web/t100sb07_1?co_id=2330');
    expect(await screen.findByText('已存在')).toBeVisible();
    expect(await screen.findByText('已同步')).toBeVisible();
    expect(await screen.findByText('Google Calendar 需要重新授權')).toBeVisible();
    expect(screen.getByRole('button', { name: '連線 Google Calendar' })).toBeVisible();
    expect(api.getGoogleCalendarStatus).toHaveBeenCalled();
    expect(api.listConferences).toHaveBeenCalled();
  });

  it('conference-calendar-sync / first connection / lets the user start OAuth without importing a client JSON file', async () => {
    const api = setupApi();
    api.getGoogleCalendarStatus = vi.fn(async () => ({ status: 'not-configured' }));
    render(<App />);
    await userEvent.click(await screen.findByRole('tab', { name: '法說會' }));

    expect(await screen.findByRole('button', { name: '連線 Google Calendar' })).toBeVisible();
    expect(screen.queryByRole('button', { name: /匯入 OAuth 設定檔/ })).not.toBeInTheDocument();
    await userEvent.click(screen.getByRole('button', { name: '連線 Google Calendar' }));

    expect(api.connectGoogleCalendar).toHaveBeenCalledOnce();
    expect(api.importGoogleCredentials).not.toHaveBeenCalled();
    expect(api.syncConferences).not.toHaveBeenCalled();
  });

  it('conference-calendar-sync / first connection / explains that browser authorization is pending', async () => {
    const api = setupApi();
    vi.mocked(api.connectGoogleCalendar).mockImplementation(() => new Promise(() => undefined));
    render(<App />);
    await userEvent.click(await screen.findByRole('tab', { name: '法說會' }));
    await userEvent.click(await screen.findByRole('button', { name: '連線 Google Calendar' }));

    expect(await screen.findByRole('status')).toHaveTextContent('請在系統瀏覽器完成 Google 授權');
    expect(screen.getByRole('button', { name: '連線 Google Calendar' })).toBeDisabled();
  });

  it('conference-calendar-sync / legacy migration UI / requests the one-time migration only after an explicit click', async () => {
    const api = setupApi();
    render(<App />);
    await userEvent.click(await screen.findByRole('tab', { name: '法說會' }));
    await screen.findByRole('heading', { name: '法說會與 Google Calendar' });
    await userEvent.click(screen.getByRole('button', { name: '匯入舊版 token（不刪除原檔）' }));
    expect(api.migrateLegacyGoogleToken).toHaveBeenCalledOnce();
  });

  it('searches and filters the saved watchlist', async () => {
    setupApi();
    const user = userEvent.setup();
    render(<App />);
    expect(await screen.findByText('台積電')).toBeVisible();
    await user.type(screen.getByRole('searchbox', { name: '搜尋關注公司' }), '2330');
    expect(await screen.findByText('台積電')).toBeVisible();
    await user.clear(screen.getByRole('searchbox', { name: '搜尋關注公司' }));
    await user.click(screen.getByRole('checkbox', { name: '只顯示啟用公司' }));
    expect(await screen.findByText('台積電')).toBeVisible();
  });

  it('shows failure reason distinctly from empty result and reports the latest schedule status', async () => {
    const api = setupApi();
    render(<App />);
    await userEvent.click(screen.getByRole('tab', { name: '設定' }));
    expect(await screen.findByRole('heading', { name: '監控排程' })).toBeVisible();
    expect(screen.getByRole('button', { name: '立即檢查重大訊息' })).toBeEnabled();
    expect(screen.getByRole('alert')).toHaveTextContent('MOPS 暫時無法連線');
    expect(api.getScheduleStatus).toHaveBeenCalled();
  });

  it('saves schedule settings and reports a busy result without starting another job', async () => {
    const api = setupApi();
    vi.mocked(api.runScheduledCheck).mockResolvedValue({ status: 'already-running' });
    render(<App />);
    await userEvent.click(screen.getByRole('tab', { name: '設定' }));
    await screen.findByRole('button', { name: '儲存排程' });
    await userEvent.selectOptions(screen.getByRole('combobox', { name: '重大訊息檢查頻率' }), '30');
    await userEvent.click(screen.getByRole('button', { name: '儲存排程' }));
    await waitFor(() => expect(api.saveScheduleSettings).toHaveBeenCalledWith(expect.objectContaining({ monitoring: expect.objectContaining({ intervalMinutes: 30 }) })));
    await userEvent.click(screen.getByRole('button', { name: '立即檢查重大訊息' }));
    expect(await screen.findByText('此工作目前仍在執行，未重複啟動。')).toBeVisible();
    expect(api.runScheduledCheck).toHaveBeenCalledWith('material');
  });

  it('shows a readable error when an immediate monitor check fails', async () => {
    const api = setupApi();
    vi.mocked(api.runScheduledCheck).mockRejectedValue(new Error('MOPS 請求逾時'));
    render(<App />);
    await userEvent.click(screen.getByRole('tab', { name: '設定' }));
    await screen.findByRole('button', { name: '立即檢查重大訊息' });
    await userEvent.click(screen.getByRole('button', { name: '立即檢查重大訊息' }));
    expect(await screen.findByText('MOPS 請求逾時')).toBeVisible();
  });

  it('does not call a completed manual check successful when the source status is failed or stale', async () => {
    const api = setupApi();
    const initial = await api.getScheduleStatus() as { disclosure: Record<string, unknown> };
    api.getScheduleStatus = vi.fn()
      .mockResolvedValueOnce(initial)
      .mockResolvedValueOnce({ ...initial, disclosure: { ...initial.disclosure, status: 'failed', errorMessage: 'TPEX: 回應格式無效' } })
      .mockResolvedValue({ ...initial, disclosure: { ...initial.disclosure, status: 'stale', errorMessage: null } });
    render(<App />);
    await userEvent.click(screen.getByRole('tab', { name: '設定' }));
    const button = await screen.findByRole('button', { name: '立即檢查違約交割' });
    await userEvent.click(button);
    expect(await screen.findByText('違約交割檢查失敗：TPEX: 回應格式無效')).toBeVisible();
    expect(screen.queryByText('違約交割檢查完成，請查看最新狀態。')).not.toBeInTheDocument();
    await userEvent.click(button);
    expect(await screen.findByText('違約交割官方資料尚未更新。')).toBeVisible();
  });

  it('shows that isolated mode cannot run live checks instead of offering ineffective buttons', async () => {
    const api = setupApi();
    const status = await api.getScheduleStatus();
    api.getScheduleStatus = vi.fn(async () => ({ ...(status as object), available: false }));
    render(<App />);
    await userEvent.click(screen.getByRole('tab', { name: '設定' }));
    expect(await screen.findByText('隔離測試模式不會連線官方網站；立即檢查已停用。')).toBeVisible();
    expect(screen.getByLabelText('背景執行狀態')).toHaveTextContent('隔離測試模式：背景監控已停用。');
    expect(screen.getAllByText('預計下次：隔離測試模式不執行')).toHaveLength(2);
    expect(screen.getByRole('button', { name: '立即檢查重大訊息' })).toBeDisabled();
    expect(screen.getByRole('button', { name: '立即檢查違約交割' })).toBeDisabled();
    expect(api.runScheduledCheck).not.toHaveBeenCalled();
  });

  it('shows immediate progress beside the selected job while an actual check is pending', async () => {
    const api = setupApi();
    api.runScheduledCheck = vi.fn(() => new Promise(() => undefined));
    render(<App />);
    await userEvent.click(screen.getByRole('tab', { name: '設定' }));
    await userEvent.click(await screen.findByRole('button', { name: '立即檢查重大訊息' }));
    const card = screen.getByRole('heading', { name: '重大訊息' }).closest('article')!;
    expect(within(card).getByRole('status')).toHaveTextContent('正在檢查重大訊息');
    expect(within(card).getByRole('button', { name: '檢查重大訊息中…' })).toBeDisabled();
  });

  it('reports an invalid code and adds a valid company from the keyboard', async () => {
    const api = setupApi();
    const user = userEvent.setup();
    render(<App />);
    const form = screen.getByRole('form', { name: '新增關注公司' });
    const code = within(form).getByRole('textbox', { name: '股票代號' });
    await user.type(code, 'bad');
    await user.keyboard('{Enter}');
    expect(await screen.findByRole('alert')).toHaveTextContent('股票代號必須是 4 至 6 位數字');
    expect(api.addWatchlist).not.toHaveBeenCalled();

    await user.clear(code);
    await user.type(code, '2317');
    await user.keyboard('{Enter}');
    await waitFor(() => expect(api.addWatchlist).toHaveBeenCalledWith(expect.objectContaining({
      market: 'TWSE', stockCode: '2317',
    })));
  });

  it('edits category and notes and preserves the company in the local list when disabling it', async () => {
    const api = setupApi();
    const user = userEvent.setup();
    render(<App />);
    await screen.findByText('台積電');
    const row = screen.getByRole('listitem');
    await user.clear(within(row).getByRole('textbox', { name: '分類台積電' }));
    await user.type(within(row).getByRole('textbox', { name: '分類台積電' }), '半導體');
    await user.clear(within(row).getByRole('textbox', { name: '備註台積電' }));
    await user.type(within(row).getByRole('textbox', { name: '備註台積電' }), '注意法說會');
    await user.click(within(row).getByRole('button', { name: '儲存台積電' }));
    await waitFor(() => expect(api.updateWatchlist).toHaveBeenCalledWith({ companyId: 'company-1', category: '半導體', notes: '注意法說會' }));
    await user.click(within(row).getByRole('button', { name: '停用台積電' }));
    expect(await within(row).findByText('已停用')).toBeVisible();
    expect(api.setWatchlistActive).toHaveBeenCalledWith({ companyId: 'company-1', active: false });
  });

  it('previews an old configuration before applying the import', async () => {
    const api = setupApi();
    const user = userEvent.setup();
    render(<App />);
    const config = screen.getByRole('textbox', { name: '舊設定 JSON' });
    fireEvent.change(config, { target: { value: '{"StockNumbers":"2330,2317"}' } });
    await user.click(screen.getByRole('button', { name: '預覽匯入' }));
    expect(await screen.findByText('新增 1、略過 1、失敗 1')).toBeVisible();
    expect(api.applyWatchlistImport).not.toHaveBeenCalled();
    await user.click(screen.getByRole('button', { name: '確認並匯入' }));
    await waitFor(() => expect(api.applyWatchlistImport).toHaveBeenCalledWith('{"StockNumbers":"2330,2317"}'));
  });

  it('shows every market disclosure while visually highlighting the watched company', async () => {
    const api = setupApi();
    const user = userEvent.setup();
    render(<App />);
    await user.click(screen.getByRole('tab', { name: '違約交割' }));
    expect(await screen.findByText('關注半導體')).toBeVisible();
    expect(await screen.findByText('一般公司')).toBeVisible();
    expect(screen.getByText('關注公司', { selector: '.watched-badge' })).toBeVisible();
    expect(api.listDisclosures).toHaveBeenCalledWith(expect.objectContaining({ disclosureDate: expect.any(String) }));
  });

  it('shows stale and failed disclosure sources separately from a zero-row list', async () => {
    const api = setupApi();
    vi.mocked(api.listDisclosures).mockResolvedValue([]);
    render(<App />);
    await userEvent.click(screen.getByRole('tab', { name: '違約交割' }));
    expect(await screen.findByRole('status', { name: '違約交割來源狀態' })).toHaveTextContent('資料來源部分降級');
    expect(screen.getByRole('status', { name: '違約交割來源狀態' })).toHaveTextContent('TWSE：官方尚未更新（資料日期 2026-09-25）');
    expect(screen.getByRole('status', { name: '違約交割來源狀態' })).toHaveTextContent(/預計重試 2026\/09\/26\s*19:00/);
    expect(screen.getByRole('status', { name: '違約交割來源狀態' })).not.toHaveTextContent('T19:00:00+08:00');
    expect(screen.getByRole('status', { name: '違約交割來源狀態' })).toHaveTextContent('TPEX：來源失敗');
    expect(screen.getByText('來源檢查失敗或官方資料尚未更新，不能判定為零筆。')).toBeVisible();
    expect(screen.getByRole('heading', { name: '全市場違約交割' }).closest('.section-heading')?.querySelector('.count'))
      .toHaveTextContent('未確認');
    expect(screen.queryByText('當日無個股達違約資訊揭露標準')).not.toBeInTheDocument();
    await userEvent.click(screen.getByRole('tab', { name: '總覽' }));
    expect(screen.getByRole('heading', { name: '今日違約揭露' }).closest('.overview-card'))
      .toHaveTextContent('未確認');
  });

  it('labels a current-date zero-result as no disclosure rather than an error', async () => {
    const api = setupApi();
    vi.mocked(api.listDisclosures).mockResolvedValue([]);
    vi.mocked(api.getDisclosureMonitorStatus).mockResolvedValue({
      status: 'complete', finishedAt: '2026-09-26T11:00:00.000Z', errorMessage: null,
      markets: { TWSE: { status: 'empty_success', dataDate: '2026-09-26', recordCount: 0, retryAt: null } },
    });
    render(<App />);
    await userEvent.click(screen.getByRole('tab', { name: '違約交割' }));
    fireEvent.change(screen.getByLabelText('違約交割資料日期'), { target: { value: '2026-09-26' } });
    fireEvent.change(screen.getByRole('combobox', { name: '違約交割市場' }), { target: { value: 'TWSE' } });
    expect(await screen.findByRole('status', { name: '違約交割來源狀態' })).toHaveTextContent('TWSE：本日無揭露');
    expect(screen.getByText('當日無個股達違約資訊揭露標準')).toBeVisible();
  });

  it('opens full event details without losing the search and links a correction to its original', async () => {
    const api = setupApi();
    const user = userEvent.setup();
    render(<App />);
    await user.click(screen.getByRole('tab', { name: '重大訊息' }));
    const search = screen.getByRole('searchbox', { name: '搜尋重大訊息' });
    await user.type(search, '更正');
    await user.click(await screen.findByRole('button', { name: /營收更正公告/ }));
    const expandedRow = screen.getByRole('button', { name: /營收更正公告/ }).closest('.material-row')!;
    expect(within(expandedRow as HTMLElement).getByRole('complementary', { name: '重大訊息完整內容' })).toBeVisible();
    expect(await screen.findByText('更正後完整公告內容')).toBeVisible();
    expect(screen.getByRole('link', { name: '開啟來源公告' })).toHaveAttribute('href', 'https://mops.example/event-2');
    expect(screen.getByText('原公告：營運報告')).toBeVisible();
    expect(screen.getByRole('searchbox', { name: '搜尋重大訊息' })).toHaveValue('更正');
    expect(api.getMaterialEvent).toHaveBeenCalledWith('event-2');
  });

  it('defaults to watched announcements, offers all saved announcements, and expands only the clicked row', async () => {
    const api = setupApi();
    const user = userEvent.setup();
    render(<App />);
    await user.click(screen.getByRole('tab', { name: '重大訊息' }));
    await waitFor(() => expect(api.listMaterialEvents).toHaveBeenLastCalledWith(expect.objectContaining({ watchedOnly: true })));
    expect(await screen.findByRole('button', { name: /台積電 營收更正公告/ })).toBeVisible();
    expect(screen.getByText('關注 2 筆')).toBeVisible();
    expect(screen.queryByRole('button', { name: /其他公司 其他重大公告/ })).not.toBeInTheDocument();
    await user.click(screen.getByRole('checkbox', { name: '顯示全部公告' }));
    expect(await screen.findByRole('button', { name: /其他公司 其他重大公告/ })).toBeVisible();
    expect(screen.getByText('全部 3 筆')).toBeVisible();
    await user.click(screen.getByRole('button', { name: /台積電 營收更正公告/ }));
    expect(within(screen.getByRole('button', { name: /台積電 營收更正公告/ }).closest('.material-row') as HTMLElement)
      .getByRole('complementary', { name: '重大訊息完整內容' })).toBeVisible();
    await user.click(screen.getByRole('button', { name: /其他公司 其他重大公告/ }));
    expect(within(screen.getByRole('button', { name: /其他公司 其他重大公告/ }).closest('.material-row') as HTMLElement)
      .getByRole('complementary', { name: '重大訊息完整內容' })).toBeVisible();
    expect(within(screen.getByRole('button', { name: /台積電 營收更正公告/ }).closest('.material-row') as HTMLElement)
      .queryByRole('complementary')).not.toBeInTheDocument();
  });

  it('stops loading and reports a detail retrieval error in the clicked row', async () => {
    const api = setupApi();
    vi.mocked(api.getMaterialEvent).mockRejectedValueOnce(new Error('公告詳情暫時無法取得'));
    const user = userEvent.setup();
    render(<App />);
    await user.click(screen.getByRole('tab', { name: '重大訊息' }));
    await user.click(await screen.findByRole('button', { name: /台積電 營收更正公告/ }));
    expect(await screen.findByRole('alert')).toHaveTextContent('公告詳情暫時無法取得');
    expect(screen.queryByText('載入完整公告…')).not.toBeInTheDocument();
    expect(screen.getByRole('button', { name: /台積電 營收更正公告/ })).toHaveAttribute('aria-expanded', 'false');
  });

  it('shows a source failure separately from an empty material-event result', async () => {
    const api = setupApi();
    render(<App />);
    vi.mocked(api.listMaterialEvents).mockResolvedValue([]);
    await userEvent.click(screen.getByRole('tab', { name: '重大訊息' }));
    expect(await screen.findByRole('status', { name: '重大訊息來源狀態' })).toHaveTextContent('資料來源部分降級');
    expect(screen.getByRole('status', { name: '重大訊息來源狀態' })).toHaveTextContent('MOPS 暫時無法連線');
    expect(screen.getByText('目前沒有符合條件的關注公司重大訊息；可勾選「顯示全部公告」。')).toBeVisible();
    expect(screen.getByRole('status', { name: '重大訊息來源狀態' })).toHaveTextContent('上市結構化每日對帳：資料完整');
  });

  it('opens a material event detail with keyboard activation', async () => {
    const api = setupApi();
    render(<App />);
    await userEvent.click(screen.getByRole('tab', { name: '重大訊息' }));
    const eventButton = await screen.findByRole('button', { name: /台積電 營收更正公告/ });
    eventButton.focus();
    await userEvent.keyboard('{Enter}');
    expect(await screen.findByRole('complementary', { name: '重大訊息完整內容' })).toHaveTextContent('更正後完整公告內容');
    expect(api.getMaterialEvent).toHaveBeenCalledWith('event-2');
  });

  it('routes a source-failure toast to the material source status without requesting an event', async () => {
    const api = setupApi();
    vi.mocked(api.listMaterialEvents).mockResolvedValue([]);
    render(<App />);
    (window as Window & { sendReporterNotificationRoute: (route: { type: 'source-status' }) => void })
      .sendReporterNotificationRoute({ type: 'source-status' });
    expect(await screen.findByRole('status', { name: '重大訊息來源狀態' })).toHaveTextContent('資料來源部分降級');
    expect(api.getMaterialEvent).not.toHaveBeenCalled();
  });

  it('routes a single-event toast directly to the full event details', async () => {
    const api = setupApi();
    render(<App />);
    (window as Window & { sendReporterNotificationRoute: (route: { type: 'event-detail'; eventId: string }) => void })
      .sendReporterNotificationRoute({ type: 'event-detail', eventId: 'event-2' });
    expect(await screen.findByRole('complementary', { name: '重大訊息完整內容' })).toHaveTextContent('更正後完整公告內容');
    expect(api.getMaterialEvent).toHaveBeenCalledWith('event-2');
  });

  it('opens a filtered material list from a notification batch route', async () => {
    const api = setupApi();
    render(<App />);
    await userEvent.click(screen.getByRole('tab', { name: '重大訊息' }));
    await screen.findByRole('button', { name: /台積電 營收更正公告/ });
    (window as Window & { sendReporterNotificationRoute: (route: { type: 'event-list'; eventIds: string[] }) => void })
      .sendReporterNotificationRoute({ type: 'event-list', eventIds: ['event-2'] });
    expect(await screen.findByRole('button', { name: /台積電 營收更正公告/ })).toBeVisible();
    await waitFor(() => expect(api.listMaterialEvents).toHaveBeenLastCalledWith(expect.objectContaining({ eventIds: ['event-2'] })));
    expect(screen.queryByRole('button', { name: /台積電 營運報告/ })).not.toBeInTheDocument();
  });
});

describe('desktop-app-lifecycle / login startup UI', () => {
  afterEach(() => cleanup());

  it('loads, enables and saves login launch and hidden-start preferences', async () => {
    const api = setupApi();
    const user = userEvent.setup();
    render(<App />);
    await user.click(screen.getByRole('tab', { name: '設定' }));
    expect(await screen.findByLabelText('登入 Windows 時啟動')).toBeVisible();
    await user.click(screen.getByLabelText('登入 Windows 時啟動'));
    await user.click(screen.getByLabelText('啟動後縮到系統匣'));
    await user.click(screen.getByRole('button', { name: '儲存登入啟動設定' }));
    await waitFor(() => expect(api.saveLoginStartupSettings).toHaveBeenCalledWith({ enabled: true, startHidden: true }));
  });
});
