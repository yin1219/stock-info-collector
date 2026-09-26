import { cleanup, fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
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
  }];
  let notificationRouteListener: ((route: { type: 'event-detail'; eventId: string } | { type: 'event-list'; eventIds: string[] } | { type: 'disclosure-list' }) => void) | undefined;
  const api: ReporterApi = {
    getVersion: () => 'test',
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
    removeWatchlist: vi.fn(async () => undefined),
    previewWatchlistImport: vi.fn(async () => ({ summary: { added: 1, skipped: 1, failed: 1 }, items: [] })),
    applyWatchlistImport: vi.fn(async () => ({ summary: { added: 1, skipped: 1, failed: 1 }, items: [] })),
    listDisclosures: vi.fn(async () => [
      { id: 'd1', market: 'TWSE', stockCode: '2330', companyName: '關注半導體', disclosureDate: '2026-09-26', brokerCode: 'B001', isWatched: true, contentJson: '{"amount":50000}' },
      { id: 'd2', market: 'TPEX', stockCode: '6488', companyName: '一般公司', disclosureDate: '2026-09-26', brokerCode: 'B003', isWatched: false, contentJson: '{"amount":10000}' },
    ]),
    listMaterialEvents: vi.fn(async (filter = {}) => materialEvents.filter((event) =>
      (!filter.unreadOnly || !event.readAt)
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
  beforeEach(() => vi.restoreAllMocks());

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

  it('opens full event details without losing the search and links a correction to its original', async () => {
    const api = setupApi();
    const user = userEvent.setup();
    render(<App />);
    await user.click(screen.getByRole('tab', { name: '重大訊息' }));
    const search = screen.getByRole('searchbox', { name: '搜尋重大訊息' });
    await user.type(search, '更正');
    await user.click(await screen.findByRole('button', { name: /營收更正公告/ }));
    expect(await screen.findByText('更正後完整公告內容')).toBeVisible();
    expect(screen.getByRole('link', { name: '開啟來源公告' })).toHaveAttribute('href', 'https://mops.example/event-2');
    expect(screen.getByText('原公告：營運報告')).toBeVisible();
    expect(screen.getByRole('searchbox', { name: '搜尋重大訊息' })).toHaveValue('更正');
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
