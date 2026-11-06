import {
  cleanup,
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import axe from "axe-core";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ReporterApi } from "../../src/preload/api";
import { App } from "../../src/renderer/App";

const company = {
  id: "watch-1",
  companyId: "company-1",
  market: "TWSE" as const,
  stockCode: "2330",
  name: "台積電",
  active: true,
  category: "半導體",
  notes: "追蹤法說會",
  createdAt: "2026-01-01T00:00:00.000Z",
  updatedAt: "2026-01-01T00:00:00.000Z",
};

describe("approved UI production integration", () => {
  it("opens initial setup and navigates straight to notification settings", async () => {
    localStorage.clear();
    const api = setupApi();
    api.getMarketRecords = vi.fn(async () => ({ coverage: { start: null, end: null } }));
    render(<App />);
    fireEvent.click(screen.getByRole("button", { name: "開始設定", exact: true }));
    const checklist = await screen.findByRole("dialog", { name: "首次使用設定" });
    fireEvent.click(within(checklist).getByRole("button", { name: "前往測試通知" }));
    expect(await screen.findByRole("tab", { name: "通知", exact: true })).toHaveAttribute("aria-selected", "true");
    expect(screen.queryByRole("dialog", { name: "首次使用設定" })).toBeNull();
    fireEvent.click(screen.getByRole("button", { name: "使用指南", exact: true }));
    fireEvent.click(screen.getByRole("button", { name: "首次使用設定", exact: true }));
    expect(screen.getByRole("dialog", { name: "首次使用設定" })).toBeVisible();
  });

  afterEach(cleanup);
  it("reports partial Calendar failures instead of claiming a completed sync", async () => {
    const api = setupApi();
    api.getGoogleCalendarStatus = vi.fn(async () => ({ status: "connected" }));
    api.getConferenceTaskStatus = vi.fn(async () => ({ kind: "v2" }));
    api.syncConferences = vi.fn(async () => ({
      status: "partial",
      results: [{ status: "failed" }],
    }));
    render(<App />);
    fireEvent.click(screen.getByRole("tab", { name: "法說會", exact: true }));
    const button = await screen.findByRole("button", {
      name: "立即擷取並同步法說會",
    });
    await waitFor(() => expect(button).toBeEnabled());
    fireEvent.click(button);
    expect(await screen.findByText(/失敗 1/)).toBeVisible();
    expect(screen.queryByText(/法說會同步完成/)).not.toBeInTheDocument();
  });
  it("resets disclosure filters together with a notification query", async () => {
    const api = setupApi();
    render(<App />);
    fireEvent.click(screen.getByRole("tab", { name: "違約交割", exact: true }));
    const date = screen.getByLabelText("違約交割資料日期") as HTMLInputElement;
    const today = date.value;
    fireEvent.change(date, { target: { value: "2020-01-02" } });
    fireEvent.change(screen.getByLabelText("違約交割市場"), {
      target: { value: "TPEX" },
    });
    (
      window as unknown as {
        sendReporterNotificationRoute: (route: {
          type: "disclosure-list";
        }) => void;
      }
    ).sendReporterNotificationRoute({ type: "disclosure-list" });
    await waitFor(() => expect(date).toHaveValue(today));
    expect(screen.getByLabelText("違約交割市場")).toHaveValue("all");
    expect(api.listDisclosures).toHaveBeenLastCalledWith({
      disclosureDate: today,
    });
  });
  it("opens the actual overview with seven pages, product version and no demo content", async () => {
    setupApi();
    render(<App />);
    expect(
      await screen.findByRole("heading", {
        level: 1,
        name: "總覽",
        exact: true,
      }),
    ).toBeVisible();
    expect(await screen.findByText("v2 · 版本 1.0.13")).toBeVisible();
    expect(screen.queryByText("設計預覽 · 示意資料")).not.toBeInTheDocument();
    expect(screen.queryByText("+3,281.06")).not.toBeInTheDocument();
  });
  it("edits saved companies through a dialog and opens the add form only on request", async () => {
    const api = setupApi();
    render(<App />);
    fireEvent.click(screen.getByRole("tab", { name: "關注公司", exact: true }));
    fireEvent.click(
      await screen.findByRole("button", { name: "編輯台積電", exact: true }),
    );
    const dialog = screen.getByRole("dialog");
    fireEvent.change(within(dialog).getByLabelText("分類"), {
      target: { value: "新分類" },
    });
    fireEvent.click(
      within(dialog).getByRole("button", { name: "儲存", exact: true }),
    );
    await waitFor(() =>
      expect(api.updateWatchlist).toHaveBeenCalledWith(
        expect.objectContaining({ companyId: "company-1", category: "新分類" }),
      ),
    );
    expect(await screen.findByText("新分類")).toBeVisible();
    fireEvent.click(
      screen.getByRole("button", { name: "新增關注公司", exact: true }),
    );
    expect(
      within(screen.getByRole("dialog")).getByLabelText("股票代號"),
    ).toBeVisible();
  });
  it("organizes real settings into four tabs and sends notifications through the preload API", async () => {
    const api = setupApi();
    render(<App />);
    fireEvent.click(screen.getByRole("tab", { name: "設定", exact: true }));
    expect(
      await screen.findByRole("heading", {
        level: 1,
        name: "設定",
        exact: true,
      }),
    ).toBeVisible();
    const tabs = screen.getByRole("tablist", { name: "設定分類" });
    expect(
      within(tabs)
        .getAllByRole("tab")
        .map((el) => el.textContent),
    ).toEqual(["監控排程", "通知", "啟動與系統匣", "資料管理"]);
    fireEvent.click(
      within(tabs).getByRole("tab", { name: "通知", exact: true }),
    );
    fireEvent.click(
      await screen.findByRole("button", { name: "發送測試通知", exact: true }),
    );
    await waitFor(() =>
      expect(api.sendTestNotification).toHaveBeenCalledOnce(),
    );
    expect(
      await screen.findByText("隔離測試模式不顯示 Windows 通知。"),
    ).toBeVisible();
  });
});

function setupApi() {
  let rows = [company];
  let materialEvents = [
    {
      id: "event-2",
      companyId: "company-1",
      market: "TWSE" as const,
      stockCode: "2330",
      companyName: "台積電",
      title: "營收更正公告",
      content: "更正後完整公告內容",
      source: "mops",
      sourceUrl: "https://mops.example/event-2",
      publishedAt: "2026-09-26T09:30:00.000Z",
      discoveredAt: "2026-09-26T09:35:00.000Z",
      eventType: "correction",
      revisionOf: "event-1",
      readAt: null,
    },
    {
      id: "event-1",
      companyId: "company-1",
      market: "TWSE" as const,
      stockCode: "2330",
      companyName: "台積電",
      title: "營運報告",
      content: "原公告完整內容",
      source: "mops",
      sourceUrl: "https://mops.example/event-1",
      publishedAt: "2026-09-25T09:30:00.000Z",
      discoveredAt: "2026-09-25T09:35:00.000Z",
      eventType: "announcement",
      revisionOf: null,
      readAt: "2026-09-25T10:00:00.000Z",
    },
    {
      id: "event-3",
      companyId: "company-3",
      market: "TPEX" as const,
      stockCode: "6488",
      companyName: "其他公司",
      title: "其他重大公告",
      content: "未關注公司公告內容",
      source: "mops",
      sourceUrl: "https://mops.example/event-3",
      publishedAt: "2026-09-26T08:30:00.000Z",
      discoveredAt: "2026-09-26T08:35:00.000Z",
      eventType: "announcement",
      revisionOf: null,
      readAt: null,
    },
  ];
  let notificationRouteListener:
    | ((
        route:
          | { type: "event-detail"; eventId: string }
          | { type: "event-list"; eventIds: string[] }
          | { type: "source-status" }
          | { type: "disclosure-list" }
          | { type: "test-notification" },
      ) => void)
    | undefined;
  let scheduleSettings = {
    monitoring: {
      enabled: true,
      start: "00:00",
      end: "00:00",
      intervalMinutes: 60 as const,
    },
    disclosure: { enabled: true, runAt: "18:30" },
    notifyEmptyDefaultDisclosures: true,
  };
  let loginStartupSettings = { enabled: false, startHidden: false };
  let googleCalendarStatus = { status: "disconnected" as const };
  let conferenceTaskStatus = {
    kind: "legacy" as const,
    configPath: "F:\\ReporterV1\\config\\default.json",
    setupReady: false,
  };
  let legacyConfigReview = {
    configPath: "F:\\ReporterV1\\config\\default.json",
    hash: "a".repeat(64),
    currentCodes: ["2330"],
    activeCodes: ["2454"],
    confirmed: false,
  };
  const conferences = [
    {
      id: "conf-pending",
      companyId: "company-1",
      stockCode: "2330",
      companyName: "台積電",
      startsAt: "2026-10-01T02:00:00.000Z",
      location: "台北",
      content: "法說會",
      sourceUrl: "https://mops.twse.com.tw/mops/web/t100sb07_1?co_id=2330",
      syncStatus: "pending",
      lastError: null,
    },
    {
      id: "conf-existing",
      companyId: "company-1",
      stockCode: "2330",
      companyName: "台積電",
      startsAt: "2026-10-02T02:00:00.000Z",
      location: "線上",
      content: "法說會",
      syncStatus: "existing",
      lastError: null,
    },
    {
      id: "conf-synced",
      companyId: "company-1",
      stockCode: "2330",
      companyName: "台積電",
      startsAt: "2026-10-03T02:00:00.000Z",
      location: "線上",
      content: "法說會",
      syncStatus: "synced",
      lastError: null,
    },
    {
      id: "conf-failed",
      companyId: "company-1",
      stockCode: "2330",
      companyName: "台積電",
      startsAt: "2026-10-04T02:00:00.000Z",
      location: "台北",
      content: "法說會",
      syncStatus: "failed",
      lastError: "Google Calendar 需要重新授權",
    },
  ];
  const scheduleStatus = {
    material: {
      enabled: true,
      status: "failed",
      lastSuccessAt: "2026-09-26T02:00:00.000Z",
      lastResult: "failed",
      errorMessage: "MOPS 暫時無法連線",
      nextRunAt: "2026-09-26T03:00:00.000Z",
    },
    disclosure: {
      enabled: true,
      status: "complete",
      lastSuccessAt: "2026-09-25T10:30:00.000Z",
      lastResult: "complete",
      errorMessage: null,
      nextRunAt: "2026-09-26T10:30:00.000Z",
    },
  };
  const api: ReporterApi = {
    getVersion: vi.fn(async () => "1.0.13"),
    getMaterialMonitorStatus: vi.fn(async () => ({
      status: "degraded",
      finishedAt: "2026-09-26T11:00:00.000Z",
      errorMessage: "MOPS 暫時無法連線，RSS 備援已使用。",
      sources: [
        { source: "MOPS", status: "failed" },
        { source: "MOPS-RSS", status: "complete" },
        { source: "TWSE-reconciliation", status: "complete" },
      ],
    })),
    getGoogleCalendarStatus: vi.fn(async () => googleCalendarStatus),
    importGoogleCredentials: vi.fn(async () => true),
    migrateLegacyGoogleToken: vi.fn(async () => true),
    connectGoogleCalendar: vi.fn(async () => {
      googleCalendarStatus = { status: "connected" };
      return googleCalendarStatus;
    }),
    disconnectGoogleCalendar: vi.fn(async () => {
      googleCalendarStatus = { status: "disconnected" };
      return googleCalendarStatus;
    }),
    listConferences: vi.fn(async () => conferences),
    syncConferences: vi.fn(async () => ({ status: "complete", results: [] })),
    getConferenceTaskStatus: vi.fn(async () => conferenceTaskStatus),
    ensureConferenceTask: vi.fn(async () => {
      conferenceTaskStatus = {
        kind: "legacy",
        configPath: conferenceTaskStatus.configPath,
        setupReady: false,
      };
      return conferenceTaskStatus;
    }),
    switchFromDisabledLegacy: vi.fn(async () => ({
      kind: "v2",
      setupReady: true,
    })),
    reviewLegacyConferenceConfig: vi.fn(async () => legacyConfigReview),
    confirmLegacyConferenceConfig: vi.fn(async () => {
      legacyConfigReview = {
        ...legacyConfigReview,
        currentCodes: [...legacyConfigReview.activeCodes],
        confirmed: true,
      };
      return {
        status: "applied" as const,
        stockNumbers: [...legacyConfigReview.activeCodes],
        backupPath: "F:\\ReporterV1\\config\\default.json.pre-v2.backup",
      };
    }),
    getScheduleSettings: vi.fn(async () => scheduleSettings),
    saveScheduleSettings: vi.fn(async (value) => {
      scheduleSettings = value;
      return value;
    }),
    getScheduleStatus: vi.fn(async () => scheduleStatus),
    runScheduledCheck: vi.fn(async () => ({ status: "started" })),
    sendTestNotification: vi.fn(async () => ({
      status: "suppressed" as const,
    })),
    scheduleTestNotification: vi.fn(async () => ({
      status: "scheduled" as const,
      scheduledAt: "2026-09-28T00:01:00.000Z",
    })),
    exportUserData: vi.fn(async () => ({ status: "cancelled" as const })),
    getLoginStartupSettings: vi.fn(async () => loginStartupSettings),
    saveLoginStartupSettings: vi.fn(async (value) => {
      loginStartupSettings = value;
      return value;
    }),
    listWatchlist: vi.fn(async (filter = {}) =>
      rows.filter(
        (row) =>
          (!filter.activeOnly || row.active) &&
          (!filter.query ||
            `${row.stockCode} ${row.name}`.includes(filter.query)),
      ),
    ),
    addWatchlist: vi.fn(async (input) => {
      rows = [
        { ...company, ...input, id: "watch-2", companyId: "company-2" },
        ...rows,
      ];
      return rows[0];
    }),
    updateWatchlist: vi.fn(async (input) => {
      rows = rows.map((row) =>
        row.companyId === input.companyId
          ? { ...row, category: input.category, notes: input.notes }
          : row,
      );
      return rows[0];
    }),
    setWatchlistActive: vi.fn(async ({ active }) => {
      rows = rows.map((row) => ({ ...row, active }));
      return rows[0];
    }),
    removeWatchlist: vi.fn(async ({ companyId }) => {
      rows = rows.filter((row) => row.companyId !== companyId);
    }),
    previewWatchlistImport: vi.fn(async () => ({
      summary: { added: 1, skipped: 1, failed: 1 },
      items: [],
    })),
    applyWatchlistImport: vi.fn(async () => ({
      summary: { added: 1, skipped: 1, failed: 1 },
      items: [],
    })),
    listDisclosures: vi.fn(async () => [
      {
        id: "d1",
        market: "TWSE",
        stockCode: "2330",
        companyName: "關注半導體",
        disclosureDate: "2026-09-26",
        brokerCode: "B001",
        isWatched: true,
        contentJson: '{"amount":50000}',
      },
      {
        id: "d2",
        market: "TPEX",
        stockCode: "6488",
        companyName: "一般公司",
        disclosureDate: "2026-09-26",
        brokerCode: "B003",
        isWatched: false,
        contentJson: '{"amount":10000}',
      },
    ]),
    getDisclosureMonitorStatus: vi.fn(async () => ({
      status: "degraded",
      finishedAt: "2026-09-26T11:00:00.000Z",
      errorMessage: "TPEX: 來源逾時",
      markets: {
        TWSE: {
          status: "stale",
          dataDate: "2026-09-25",
          recordCount: 3,
          retryAt: "2026-09-26T19:00:00+08:00",
        },
        TPEX: {
          status: "failed",
          dataDate: null,
          recordCount: 0,
          retryAt: null,
          errorMessage: "來源逾時",
        },
      },
    })),
    listMaterialEvents: vi.fn(async (filter = {}) =>
      materialEvents.filter(
        (event) =>
          (!filter.unreadOnly || !event.readAt) &&
          (!filter.watchedOnly ||
            rows.some(
              (row) => row.companyId === event.companyId && row.active,
            )) &&
          (!filter.eventIds || filter.eventIds.includes(event.id)) &&
          (!filter.query ||
            `${event.stockCode} ${event.companyName} ${event.title} ${event.content}`.includes(
              filter.query,
            )),
      ),
    ),
    onNotificationRoute: vi.fn((listener) => {
      notificationRouteListener = listener;
      return () => {
        notificationRouteListener = undefined;
      };
    }),
    getMaterialEvent: vi.fn(async (id) => {
      const event = materialEvents.find((item) => item.id === id)!;
      return {
        ...event,
        relatedRevisions: [],
        original: event.revisionOf
          ? materialEvents.find((item) => item.id === event.revisionOf)
          : null,
      };
    }),
    markMaterialEventRead: vi.fn(async (id) => {
      const event = materialEvents.find((item) => item.id === id)!;
      event.readAt = "2026-09-26T10:00:00.000Z";
      return event;
    }),
    softDeleteMaterialEvent: vi.fn(async (id) => {
      materialEvents = materialEvents.filter((event) => event.id !== id);
    }),
  };
  Object.defineProperty(window, "reporterApi", {
    configurable: true,
    value: api,
  });
  Object.defineProperty(window, "sendReporterNotificationRoute", {
    configurable: true,
    value: (
      route: Parameters<NonNullable<typeof notificationRouteListener>>[0],
    ) => notificationRouteListener?.(route),
  });
  return api;
}

function renderWatchlist() {
  const result = render(<App />);
  fireEvent.click(screen.getByRole("tab", { name: "關注公司", exact: true }));
  return result;
}

describe("watchlist-management / UI", () => {
  it("desktop-app-lifecycle / installed version UI / displays the running patch version in the sidebar", async () => {
    const api = setupApi();
    renderWatchlist();
    expect(await screen.findByText("v2 · 版本 1.0.13")).toBeVisible();
    expect(api.getVersion).toHaveBeenCalledOnce();
  });
  afterEach(() => cleanup());

  it("removes a company from the visible list instead of showing it as disabled", async () => {
    const api = setupApi();
    const user = userEvent.setup();
    renderWatchlist();
    await user.click(screen.getByRole("tab", { name: "關注公司" }));
    expect(await screen.findByText("台積電")).toBeVisible();
    await user.click(screen.getByRole("button", { name: "移除台積電" }));
    await userEvent.click(
      await screen.findByRole("button", { name: "確認移除" }),
    );
    await waitFor(() =>
      expect(screen.queryByText("台積電")).not.toBeInTheDocument(),
    );
    expect(screen.getByText("目前沒有符合條件的公司。")).toBeVisible();
    expect(api.removeWatchlist).toHaveBeenCalledWith({
      companyId: "company-1",
    });
  });
  beforeEach(() => vi.restoreAllMocks());

  it("desktop-app-lifecycle / overview dashboard / summarizes saved watchlist and current monitor data", async () => {
    const api = setupApi();
    renderWatchlist();
    expect(
      screen.getAllByRole("tab").map((tab) => tab.textContent?.trim()),
    ).toEqual([
      "總覽",
      "關注公司",
      "重大訊息",
      "違約交割",
      "法說會",
      "盤面紀錄",
      "設定",
    ]);
    await userEvent.click(await screen.findByRole("tab", { name: "總覽" }));
    expect(await screen.findByRole("heading", { name: "總覽" })).toBeVisible();
    await waitFor(() =>
      expect(document.querySelector(".overview-summary")).toHaveTextContent(
        "關注公司 1 家",
      ),
    );
    expect(
      await screen.findByRole("heading", { name: "最新重大訊息" }),
    ).toBeVisible();
    expect(
      await screen.findByRole("heading", { name: "接下來的檢查" }),
    ).toBeVisible();
    expect(
      await screen.findByRole("heading", { name: "下一場法說會" }),
    ).toBeVisible();
    expect(screen.getByRole("button", { name: "查看來源狀態" })).toBeVisible();
    expect(api.getMaterialMonitorStatus).toHaveBeenCalled();
    expect(api.listDisclosures).toHaveBeenCalled();
    expect(api.listConferences).toHaveBeenCalled();
    expect(api.getScheduleStatus).toHaveBeenCalled();
  });

  it("desktop-app-lifecycle / overview dashboard / distinguishes an unrun monitor from a still-loading status", async () => {
    const api = setupApi();
    Object.defineProperty(api, "getMaterialMonitorStatus", {
      value: vi.fn(async () => null),
    });
    Object.defineProperty(api, "getDisclosureMonitorStatus", {
      value: vi.fn(async () => null),
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "總覽" }));

    expect(await screen.findByText("來源尚未檢查")).toBeVisible();
    expect(screen.queryByText("載入中…")).not.toBeInTheDocument();
    expect(screen.queryByText("載入中…")).not.toBeInTheDocument();
  });

  it("local-data-management / export UI / lets the user export and reports success or failure", async () => {
    const api = setupApi() as ReporterApi & {
      exportUserData: ReturnType<typeof vi.fn>;
    };
    api.exportUserData = vi
      .fn()
      .mockResolvedValueOnce({
        status: "exported",
        destination: "C:/Documents/reporter.json",
      })
      .mockRejectedValueOnce(new Error("無法寫入所選位置"));
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "設定" }));
    await userEvent.click(
      screen.getByRole("tab", { name: "資料管理", exact: true }),
    );

    await userEvent.click(
      await screen.findByRole("button", { name: "匯出本機資料" }),
    );
    expect(api.exportUserData).toHaveBeenCalledOnce();
    expect(await screen.findByRole("status")).toHaveTextContent(
      "資料已匯出至 C:/Documents/reporter.json",
    );

    await userEvent.click(screen.getByRole("button", { name: "匯出本機資料" }));
    expect(await screen.findByText("無法寫入所選位置")).toHaveAttribute(
      "role",
      "alert",
    );
  });

  it("notification-delivery / empty disclosure preference UI / saves whether zero-result notices are enabled", async () => {
    const api = setupApi();
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "設定" }));
    const preference = await screen.findByRole("checkbox", {
      name: "本日無違約揭露時也通知",
    });
    expect(preference).toBeChecked();
    await userEvent.click(preference);
    await userEvent.click(screen.getByRole("button", { name: "儲存變更" }));

    await waitFor(() =>
      expect(api.saveScheduleSettings).toHaveBeenCalledWith(
        expect.objectContaining({
          notifyEmptyDefaultDisclosures: false,
        }),
      ),
    );
  });

  it("watchlist-management / legacy schedule UI / reviews old and active codes before confirming v1 config changes", async () => {
    const api = setupApi();
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "設定" }));
    expect(
      await screen.findByRole("heading", { name: "法說會每日工作" }),
    ).toBeVisible();
    expect(await screen.findByText(/沿用既有 v1 工作/)).toBeVisible();
    await userEvent.click(screen.getByRole("button", { name: "核對舊版股號" }));
    expect(await screen.findByText(/舊版目前：2330/)).toBeVisible();
    expect(screen.getByText(/v2 啟用：2454/)).toBeVisible();
    expect(api.confirmLegacyConferenceConfig).not.toHaveBeenCalled();
    await userEvent.click(
      screen.getByRole("button", { name: "確認套用到舊版工作" }),
    );
    await waitFor(() =>
      expect(api.confirmLegacyConferenceConfig).toHaveBeenCalledWith(
        "a".repeat(64),
      ),
    );
    expect(await screen.findByText(/已套用到舊版/)).toBeVisible();
  });

  it("watchlist-management / disabled legacy schedule UI / still offers reviewed stock-code sync for rollback", async () => {
    const api = setupApi();
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "legacy-disabled",
      configPath: "F:\\ReporterV1\\config\\default.json",
      setupReady: true,
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "設定" }));
    expect(await screen.findByText(/既有 v1 工作已停用/)).toBeVisible();
    await userEvent.click(screen.getByRole("button", { name: "核對舊版股號" }));
    expect(await screen.findByText(/舊版目前：2330/)).toBeVisible();
    await userEvent.click(
      screen.getByRole("button", { name: "確認套用到舊版工作" }),
    );
    await waitFor(() =>
      expect(api.confirmLegacyConferenceConfig).toHaveBeenCalledWith(
        "a".repeat(64),
      ),
    );
  });

  it("watchlist-management / v2 schedule UI / offers reviewed stock-code sync for a retained disabled v1 task", async () => {
    const api = setupApi();
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "v2",
      legacyConfigPath: "F:\\ReporterV1\\config\\default.json",
      setupReady: true,
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "設定" }));
    expect(await screen.findByText(/v2 每日工作已建立/)).toBeVisible();
    expect(screen.getByText(/Windows 每日 19:00 執行/)).toBeVisible();
    await userEvent.click(
      await screen.findByRole("button", { name: "核對舊版股號" }),
    );
    expect(await screen.findByText(/舊版目前：2330/)).toBeVisible();
    await userEvent.click(
      screen.getByRole("button", { name: "確認套用到舊版工作" }),
    );
    await waitFor(() =>
      expect(api.confirmLegacyConferenceConfig).toHaveBeenCalledWith(
        "a".repeat(64),
      ),
    );
  });

  it("windows-distribution / schedule UI / explains and retries a verified old v2 time migration", async () => {
    const api = setupApi();
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "upgrade-v2-time",
      setupReady: true,
    });
    vi.mocked(api.ensureConferenceTask).mockResolvedValue({
      kind: "v2",
      setupReady: true,
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "設定" }));
    expect(await screen.findByText(/17:45.*19:00/)).toBeVisible();
    await userEvent.click(
      screen.getByRole("button", { name: "重試遷移至 19:00" }),
    );
    await waitFor(() =>
      expect(api.ensureConferenceTask).toHaveBeenCalledOnce(),
    );
  });

  it("watchlist-management / v2 rollback UI / reports confirmed v1 stock-code synchronization after a watchlist change", async () => {
    const api = setupApi();
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "v2",
      legacyConfigPath: "F:\\ReporterV1\\config\\default.json",
      setupReady: true,
    });
    vi.mocked(api.reviewLegacyConferenceConfig).mockResolvedValue({
      configPath: "F:\\ReporterV1\\config\\default.json",
      hash: "b".repeat(64),
      currentCodes: ["2330", "2454"],
      activeCodes: ["2330", "2454"],
      confirmed: true,
    });
    renderWatchlist();
    await userEvent.click(
      screen.getByRole("button", { name: "新增關注公司", exact: true }),
    );
    await userEvent.type(
      screen.getByRole("textbox", { name: "股票代號" }),
      "2454",
    );
    await userEvent.click(screen.getByRole("button", { name: "新增公司" }));
    expect(await screen.findByText("v1 法說會股號已套用")).toBeVisible();
    expect(api.reviewLegacyConferenceConfig).toHaveBeenCalledOnce();
  });

  it("windows-distribution / legacy task health UI / warns when the sole v1 task last returned an error", async () => {
    const api = setupApi();
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "legacy",
      configPath: "F:\\ReporterV1\\config\\default.json",
      lastTaskResult: 0x800710e0,
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "設定" }));
    expect(
      await screen.findByText(/最近一次 Windows 工作結果異常.*0x800710E0/),
    ).toHaveAttribute("role", "alert");
    expect(api.ensureConferenceTask).not.toHaveBeenCalled();
    await userEvent.click(screen.getByRole("tab", { name: "法說會" }));
    expect(
      await screen.findByText(/最近一次 Windows 工作結果異常/),
    ).toBeVisible();
  });

  it("windows-distribution / v2 task health UI / distinguishes a new task from a failed task", async () => {
    const api = setupApi();
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "v2",
      lastTaskResult: 0x41303,
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "設定" }));
    expect(await screen.findByText(/v2 法說會工作尚未執行/)).toHaveAttribute(
      "role",
      "status",
    );
    expect(screen.queryByText(/Windows 工作結果異常/)).not.toBeInTheDocument();
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "v2",
      lastTaskResult: 0x800710e0,
    });
    await userEvent.click(screen.getByRole("tab", { name: "法說會" }));
    expect(
      await screen.findByText(/最近一次 Windows 工作結果異常.*0x800710E0/),
    ).toHaveAttribute("role", "alert");
  });

  it("watchlist-management / legacy schedule UI / cannot replace the last v1 stock code with an empty active list", async () => {
    const api = setupApi();
    vi.mocked(api.reviewLegacyConferenceConfig).mockResolvedValue({
      configPath: "F:\\ReporterV1\\config\\default.json",
      hash: "a".repeat(64),
      currentCodes: ["2330"],
      activeCodes: [],
      confirmed: false,
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "設定" }));
    await userEvent.click(screen.getByRole("button", { name: "核對舊版股號" }));
    expect(await screen.findByText(/舊版目前：2330/)).toBeVisible();
    expect(
      screen.getByRole("button", { name: "確認套用到舊版工作" }),
    ).toBeDisabled();
    expect(api.confirmLegacyConferenceConfig).not.toHaveBeenCalled();
  });

  it("windows-distribution / schedule UI / never calls registration when task status is unknown", async () => {
    const api = setupApi();
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "blocked",
      reason: "無法確認既有工作",
      setupReady: false,
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "設定" }));
    expect(await screen.findByText(/無法確認既有工作/)).toBeVisible();
    expect(
      screen.queryByRole("button", { name: "建立 v2 每日工作" }),
    ).not.toBeInTheDocument();
    expect(api.ensureConferenceTask).not.toHaveBeenCalled();
  });

  it("windows-distribution / schedule UI / requires readiness and an explicit click before creating the v2 daily task", async () => {
    const api = setupApi();
    vi.mocked(api.getConferenceTaskStatus)
      .mockResolvedValueOnce({ kind: "create-v2", setupReady: false })
      .mockResolvedValueOnce({ kind: "create-v2", setupReady: false })
      .mockResolvedValueOnce({ kind: "create-v2", setupReady: true });
    vi.mocked(api.ensureConferenceTask).mockResolvedValue({
      kind: "v2",
      setupReady: true,
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "設定" }));
    const button = await screen.findByRole("button", {
      name: "建立 v2 每日工作",
    });
    expect(button).toBeDisabled();
    expect(api.ensureConferenceTask).not.toHaveBeenCalled();
    await userEvent.click(screen.getByRole("tab", { name: "法說會" }));
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    await waitFor(() =>
      expect(
        screen.getByRole("button", { name: "建立 v2 每日工作" }),
      ).toBeEnabled(),
    );
    expect(api.ensureConferenceTask).not.toHaveBeenCalled();
    await userEvent.click(
      screen.getByRole("button", { name: "建立 v2 每日工作" }),
    );
    await waitFor(() =>
      expect(api.ensureConferenceTask).toHaveBeenCalledOnce(),
    );
    expect(await screen.findByText("v2 每日工作已就緒。")).toBeVisible();
  });

  it("watchlist-management / v1 coexistence UI / refreshes the changed v2 list and warns when old config was not synchronized", async () => {
    const api = setupApi();
    vi.mocked(api.removeWatchlist).mockImplementationOnce(async () => {
      vi.mocked(api.listWatchlist).mockResolvedValue([]);
      throw new Error(
        "v2 清單已更新，但 v1 法說會工作未同步：設定已由外部修改",
      );
    });
    renderWatchlist();
    await userEvent.click(
      await screen.findByRole("button", { name: "移除台積電" }),
    );
    await userEvent.click(
      await screen.findByRole("button", { name: "確認移除" }),
    );
    expect(await screen.findByRole("alert")).toHaveTextContent(
      "v2 清單已更新，但 v1 法說會工作未同步",
    );
    await waitFor(() =>
      expect(
        screen.queryByRole("dialog", { name: "移除關注公司" }),
      ).not.toBeInTheDocument(),
    );
    await waitFor(() =>
      expect(
        screen.queryByRole("button", { name: "移除台積電" }),
      ).not.toBeInTheDocument(),
    );
  });

  it("watchlist-management / v1 coexistence UI / reports a verified legacy config update after adding a company", async () => {
    const api = setupApi();
    vi.mocked(api.reviewLegacyConferenceConfig).mockResolvedValue({
      configPath: "F:\\ReporterV1\\config\\default.json",
      hash: "b".repeat(64),
      currentCodes: ["2330", "2454"],
      activeCodes: ["2330", "2454"],
      confirmed: true,
    });
    renderWatchlist();
    await userEvent.click(
      screen.getByRole("button", { name: "新增關注公司", exact: true }),
    );
    await userEvent.type(
      screen.getByRole("textbox", { name: "股票代號" }),
      "2454",
    );
    await userEvent.click(screen.getByRole("button", { name: "新增公司" }));
    expect(await screen.findByText("v1 法說會股號已套用")).toBeVisible();
    expect(api.reviewLegacyConferenceConfig).toHaveBeenCalledOnce();
  });

  it("watchlist-management / disabled v1 coexistence UI / reports confirmed rollback-list synchronization", async () => {
    const api = setupApi();
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "legacy-disabled",
      configPath: "F:\\ReporterV1\\config\\default.json",
      setupReady: true,
    });
    vi.mocked(api.reviewLegacyConferenceConfig).mockResolvedValue({
      configPath: "F:\\ReporterV1\\config\\default.json",
      hash: "b".repeat(64),
      currentCodes: ["2330", "2454"],
      activeCodes: ["2330", "2454"],
      confirmed: true,
    });
    renderWatchlist();
    await userEvent.click(
      screen.getByRole("button", { name: "新增關注公司", exact: true }),
    );
    await userEvent.type(
      screen.getByRole("textbox", { name: "股票代號" }),
      "2454",
    );
    await userEvent.click(screen.getByRole("button", { name: "新增公司" }));
    expect(await screen.findByText("v1 法說會股號已套用")).toBeVisible();
    expect(api.reviewLegacyConferenceConfig).toHaveBeenCalledOnce();
  });

  it("watchlist-management / v1 coexistence UI / never claims application when the legacy config cannot be verified", async () => {
    const api = setupApi();
    vi.mocked(api.reviewLegacyConferenceConfig).mockResolvedValue({
      configPath: "F:\\ReporterV1\\config\\default.json",
      hash: "b".repeat(64),
      currentCodes: ["2330"],
      activeCodes: ["2330", "2454"],
      confirmed: false,
    });
    renderWatchlist();
    await userEvent.click(
      screen.getByRole("button", { name: "新增關注公司", exact: true }),
    );
    await userEvent.type(
      screen.getByRole("textbox", { name: "股票代號" }),
      "2454",
    );
    await userEvent.click(screen.getByRole("button", { name: "新增公司" }));
    expect(await screen.findByRole("alert")).toHaveTextContent(
      "無法確認 v1 法說會股號已套用",
    );
    expect(screen.queryByText("v1 法說會股號已套用")).not.toBeInTheDocument();
  });

  it("desktop-app-lifecycle / accessibility / exposes labelled controls and landmarks on primary screens", async () => {
    setupApi();
    renderWatchlist();
    for (const screenName of [
      "總覽",
      "關注公司",
      "重大訊息",
      "違約交割",
      "法說會",
      "設定",
    ]) {
      await userEvent.click(screen.getByRole("tab", { name: screenName }));
      await waitFor(async () => {
        const result = await axe.run(document.body, {
          runOnly: {
            type: "tag",
            values: ["wcag2a", "wcag2aa", "wcag21a", "wcag21aa"],
          },
        });
        expect(
          result.violations.map(({ id, help }) => `${id}: ${help}`),
        ).toEqual([]);
      });
    }
  });

  it("desktop-app-lifecycle / keyboard navigation / moves through the primary navigation without a pointer", async () => {
    setupApi();
    const user = userEvent.setup();
    renderWatchlist();
    screen.getByRole("tab", { name: "總覽" }).focus();
    expect(screen.getByRole("tab", { name: "總覽" })).toHaveFocus();
    await user.tab();
    expect(screen.getByRole("tab", { name: "關注公司" })).toHaveFocus();
    await user.keyboard("{Enter}");
    expect(
      await screen.findByRole("heading", { name: "關注公司" }),
    ).toBeVisible();
  });

  it("conference-calendar-sync / UI / shows sync outcomes and reauthorization action for saved conferences", async () => {
    const api = setupApi();
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "法說會" }));
    expect(
      await screen.findByRole("heading", { name: "法說會" }),
    ).toBeVisible();
    expect(await screen.findByText("待同步")).toBeVisible();
    expect(
      screen.getByRole("link", { name: "開啟法說會來源" }),
    ).toHaveAttribute(
      "href",
      "https://mops.twse.com.tw/mops/web/t100sb07_1?co_id=2330",
    );
    expect(await screen.findByText("已存在")).toBeVisible();
    expect(await screen.findByText("已同步")).toBeVisible();
    expect(
      await screen.findByText("Google Calendar 需要重新授權"),
    ).toBeVisible();
    expect(
      screen.getByRole("button", { name: "連線 Google Calendar" }),
    ).toBeVisible();
    expect(api.getGoogleCalendarStatus).toHaveBeenCalled();
    expect(api.listConferences).toHaveBeenCalled();
  });

  it("conference-calendar-sync / UI / updates to disconnected even when remote revocation is unconfirmed", async () => {
    const api = setupApi();
    let status: "connected" | "disconnected" = "connected";
    api.getGoogleCalendarStatus = vi.fn(async () => ({ status }));
    api.disconnectGoogleCalendar = vi.fn(async () => {
      status = "disconnected";
      throw new Error("本機已中斷 Google Calendar，但 Google 端撤銷未確認");
    });
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "法說會" }));
    await userEvent.click(
      await screen.findByRole("button", { name: "中斷 Google Calendar" }),
    );

    expect(await screen.findByText("尚未連線")).toBeVisible();
    expect(
      screen.getByText(/本機已中斷 Google Calendar，但 Google 端撤銷未確認/),
    ).toBeVisible();
    expect(
      screen.getByRole("button", { name: "連線 Google Calendar" }),
    ).toBeVisible();
  });

  it("conference-calendar-sync / first connection / lets the user start OAuth without importing a client JSON file", async () => {
    const api = setupApi();
    api.getGoogleCalendarStatus = vi.fn(async () => ({
      status: "not-configured",
    }));
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "法說會" }));

    expect(
      await screen.findByRole("button", { name: "連線 Google Calendar" }),
    ).toBeVisible();
    expect(
      screen.queryByRole("button", { name: /匯入 OAuth 設定檔/ }),
    ).not.toBeInTheDocument();
    await userEvent.click(
      screen.getByRole("button", { name: "連線 Google Calendar" }),
    );

    expect(api.connectGoogleCalendar).toHaveBeenCalledOnce();
    expect(api.importGoogleCredentials).not.toHaveBeenCalled();
    expect(api.syncConferences).not.toHaveBeenCalled();
  });

  it("conference-calendar-sync / first connection / explains that browser authorization is pending", async () => {
    const api = setupApi();
    vi.mocked(api.connectGoogleCalendar).mockImplementation(
      () => new Promise(() => undefined),
    );
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "法說會" }));
    await userEvent.click(
      await screen.findByRole("button", { name: "連線 Google Calendar" }),
    );

    expect(
      await screen.findByText(
        "請在系統瀏覽器完成 Google 授權；完成後此頁會更新連線狀態。",
      ),
    ).toHaveAttribute("role", "status");
    expect(
      screen.getByRole("button", { name: "連線 Google Calendar" }),
    ).toBeDisabled();
  });

  it("conference-calendar-sync / single writer UI / explains legacy ownership and disables manual v2 sync", async () => {
    const api = setupApi();
    vi.mocked(api.getGoogleCalendarStatus).mockResolvedValue({
      status: "connected",
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "法說會" }));
    expect(
      await screen.findByText(/既有 v1 工作仍是唯一受管法說會排程/),
    ).toBeVisible();
    expect(
      screen.getByRole("button", { name: "立即擷取並同步法說會" }),
    ).toBeDisabled();
    expect(api.syncConferences).not.toHaveBeenCalled();
  });

  it("conference-calendar-sync / single writer UI / permits manual sync only for a verified v2 owner", async () => {
    const api = setupApi();
    vi.mocked(api.getGoogleCalendarStatus).mockResolvedValue({
      status: "connected",
    });
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "v2",
      setupReady: true,
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "法說會" }));
    const button = await screen.findByRole("button", {
      name: "立即擷取並同步法說會",
    });
    await waitFor(() => expect(button).toBeEnabled());
    await userEvent.click(button);
    expect(api.syncConferences).toHaveBeenCalledOnce();
  });

  it("conference-calendar-sync / disabled legacy acceptance / permits manual sync without claiming a daily schedule", async () => {
    const api = setupApi();
    vi.mocked(api.getGoogleCalendarStatus).mockResolvedValue({
      status: "connected",
    });
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "legacy-disabled",
      configPath: "F:\\ReporterV1\\config\\default.json",
      setupReady: true,
    });
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "法說會" }));
    expect(
      await screen.findByText(/v1 工作已停用.*尚無每日排程/),
    ).toBeVisible();
    const button = await screen.findByRole("button", {
      name: "立即擷取並同步法說會",
    });
    await waitFor(() => expect(button).toBeEnabled());
    await userEvent.click(button);
    expect(api.syncConferences).toHaveBeenCalledOnce();
  });

  it("windows-distribution / explicit schedule switch / requires confirmation before replacing disabled v1 ownership", async () => {
    const api = setupApi();
    vi.mocked(api.getConferenceTaskStatus).mockResolvedValue({
      kind: "legacy-disabled",
      configPath: "F:\\ReporterV1\\config\\default.json",
      setupReady: true,
    });
    const confirmation = vi
      .spyOn(window, "confirm")
      .mockReturnValueOnce(false)
      .mockReturnValueOnce(true);
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    const button = await screen.findByRole("button", {
      name: "切換為 v2 每日排程",
    });
    await userEvent.click(button);
    expect(api.switchFromDisabledLegacy).not.toHaveBeenCalled();
    await userEvent.click(button);
    await waitFor(() =>
      expect(api.switchFromDisabledLegacy).toHaveBeenCalledOnce(),
    );
    expect(await screen.findByText("v2 每日工作已就緒。")).toBeVisible();
    expect(confirmation).toHaveBeenCalledTimes(2);
    confirmation.mockRestore();
  });

  it("conference-calendar-sync / legacy migration UI / requests the one-time migration only after an explicit click", async () => {
    const api = setupApi();
    renderWatchlist();
    await userEvent.click(await screen.findByRole("tab", { name: "法說會" }));
    await screen.findByRole("heading", { name: "法說會" });
    await userEvent.click(
      screen.getByRole("button", { name: "匯入舊版 token（不刪除原檔）" }),
    );
    expect(api.migrateLegacyGoogleToken).toHaveBeenCalledOnce();
  });

  it("searches and filters the saved watchlist", async () => {
    setupApi();
    const user = userEvent.setup();
    renderWatchlist();
    expect(await screen.findByText("台積電")).toBeVisible();
    await user.type(
      screen.getByRole("searchbox", { name: "搜尋關注公司" }),
      "2330",
    );
    expect(await screen.findByText("台積電")).toBeVisible();
    await user.clear(screen.getByRole("searchbox", { name: "搜尋關注公司" }));
    await user.click(screen.getByRole("button", { name: /^監控中 / }));
    expect(await screen.findByText("台積電")).toBeVisible();
  });

  it("shows failure reason distinctly from empty result and reports the latest schedule status", async () => {
    const api = setupApi();
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    expect(await screen.findByRole("heading", { name: "設定" })).toBeVisible();
    expect(
      screen.getByRole("button", { name: "立即檢查重大訊息" }),
    ).toBeEnabled();
    expect(screen.getByRole("alert")).toHaveTextContent("MOPS 暫時無法連線");
    expect(api.getScheduleStatus).toHaveBeenCalled();
  });

  it("saves schedule settings and reports a busy result without starting another job", async () => {
    const api = setupApi();
    vi.mocked(api.runScheduledCheck).mockResolvedValue({
      status: "already-running",
    });
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    await screen.findByRole("button", { name: "儲存變更" });
    await userEvent.selectOptions(
      screen.getByRole("combobox", { name: "重大訊息檢查頻率" }),
      "30",
    );
    await userEvent.click(screen.getByRole("button", { name: "儲存變更" }));
    await waitFor(() =>
      expect(api.saveScheduleSettings).toHaveBeenCalledWith(
        expect.objectContaining({
          monitoring: expect.objectContaining({ intervalMinutes: 30 }),
        }),
      ),
    );
    await userEvent.click(
      screen.getByRole("button", { name: "立即檢查重大訊息" }),
    );
    expect(
      await screen.findByText("此工作目前仍在執行，未重複啟動。"),
    ).toBeVisible();
    expect(api.runScheduledCheck).toHaveBeenCalledWith("material");
  });

  it("shows a readable error when an immediate monitor check fails", async () => {
    const api = setupApi();
    vi.mocked(api.runScheduledCheck).mockRejectedValue(
      new Error("MOPS 請求逾時"),
    );
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    await screen.findByRole("button", { name: "立即檢查重大訊息" });
    await userEvent.click(
      screen.getByRole("button", { name: "立即檢查重大訊息" }),
    );
    expect(await screen.findByText("MOPS 請求逾時")).toBeVisible();
  });

  it("does not call a completed manual check successful when the source status is failed or stale", async () => {
    const api = setupApi();
    const initial = (await api.getScheduleStatus()) as {
      disclosure: Record<string, unknown>;
    };
    api.getScheduleStatus = vi
      .fn()
      .mockResolvedValueOnce(initial)
      .mockResolvedValueOnce(initial)
      .mockResolvedValueOnce({
        ...initial,
        disclosure: {
          ...initial.disclosure,
          status: "failed",
          errorMessage: "TPEX: 回應格式無效",
        },
      })
      .mockResolvedValue({
        ...initial,
        disclosure: {
          ...initial.disclosure,
          status: "stale",
          errorMessage: null,
        },
      });
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    const button = await screen.findByRole("button", {
      name: "立即檢查違約交割",
    });
    await userEvent.click(button);
    expect(
      await screen.findByText("違約交割檢查失敗：TPEX: 回應格式無效"),
    ).toBeVisible();
    expect(
      screen.queryByText("違約交割檢查完成，請查看最新狀態。"),
    ).not.toBeInTheDocument();
    await userEvent.click(button);
    expect(await screen.findByText("違約交割官方資料尚未更新。")).toBeVisible();
  });

  it("distinguishes a degraded material check from failure and shows the source record count", async () => {
    const api = setupApi();
    const initial = (await api.getScheduleStatus()) as {
      material: Record<string, unknown>;
    };
    api.getScheduleStatus = vi
      .fn()
      .mockResolvedValueOnce(initial)
      .mockResolvedValue({
        ...initial,
        material: {
          ...initial.material,
          status: "degraded",
          errorMessage: null,
        },
      });
    api.getMaterialMonitorStatus = vi.fn(async () => ({
      status: "degraded",
      finishedAt: "2026-09-26T11:00:00.000Z",
      errorMessage: "官網查詢可能漏筆",
      sources: [
        {
          source: "MOPS",
          status: "degraded",
          dataDate: "2026-09-26",
          recordCount: 8,
          errorMessage: null,
        },
      ],
    }));

    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    await userEvent.click(
      await screen.findByRole("button", { name: "立即檢查重大訊息" }),
    );
    expect(
      await screen.findByText(
        "重大訊息來源部分降級；資料可能不完整，請查看已保存公告與來源狀態。",
      ),
    ).toBeVisible();
    expect(
      screen.queryByText("重大訊息部分來源檢查失敗，請查看最新狀態。"),
    ).not.toBeInTheDocument();

    await userEvent.click(screen.getByRole("button", { name: "查看來源狀態" }));
    expect(
      await screen.findByRole("heading", { name: "資料來源狀態", level: 2 }),
    ).toBeVisible();
    expect(
      screen.getByText("MOPS：部分降級（資料日期 2026-09-26，來源回傳 8 筆）"),
    ).toBeVisible();
  });

  it("material-event-monitoring / reconciliation date UI / labels report dates separately from announcement dates", async () => {
    const api = setupApi();
    api.getMaterialMonitorStatus = vi.fn(async () => ({
      status: "degraded",
      finishedAt: "2026-09-30T03:00:00.000Z",
      errorMessage: null,
      sources: [
        {
          source: "MOPS",
          status: "degraded",
          dataDate: "2026-09-30",
          recordCount: 11,
          errorMessage: null,
        },
        {
          source: "TWSE-reconciliation",
          status: "complete",
          dataDate: "2026-09-30",
          recordCount: 115,
          errorMessage: null,
        },
        {
          source: "TPEX-reconciliation",
          status: "stale",
          dataDate: "2026-09-29",
          recordCount: 0,
          errorMessage: null,
        },
      ],
    }));

    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "重大訊息" }));
    await userEvent.click(
      await screen.findByRole("button", { name: "展開詳情" }),
    );
    expect(
      await screen.findByText(
        "MOPS：部分降級（資料日期 2026-09-30，來源回傳 11 筆）",
      ),
    ).toBeVisible();
    expect(
      screen.getByText(
        "上市結構化每日對帳：資料完整（出表日期 2026-09-30，來源回傳 115 筆）",
      ),
    ).toBeVisible();
    expect(
      screen.getByText(
        "上櫃結構化每日對帳：官方尚未更新（出表日期 2026-09-29，來源回傳 0 筆）",
      ),
    ).toBeVisible();
    expect(
      screen.getByText(
        "對帳的出表日期可能晚於公告發言日期；對帳筆數不等於當日新公告筆數。",
      ),
    ).toBeVisible();
    expect(
      screen.queryByText(
        "上市結構化每日對帳：資料完整（資料日期 2026-09-30，來源回傳 115 筆）",
      ),
    ).not.toBeInTheDocument();
  });

  it("shows that isolated mode cannot run live checks instead of offering ineffective buttons", async () => {
    const api = setupApi();
    const status = await api.getScheduleStatus();
    api.getScheduleStatus = vi.fn(async () => ({
      ...(status as object),
      available: false,
    }));
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    expect(
      await screen.findByText("隔離測試模式不會連線官方網站；立即檢查已停用。"),
    ).toBeVisible();
    expect(screen.getByLabelText("背景執行狀態")).toHaveTextContent(
      "隔離測試模式：背景監控已停用。",
    );
    expect(screen.getAllByText("預計下次：隔離測試模式不執行")).toHaveLength(2);
    expect(
      screen.getByRole("button", { name: "立即檢查重大訊息" }),
    ).toBeDisabled();
    expect(
      screen.getByRole("button", { name: "立即檢查違約交割" }),
    ).toBeDisabled();
    expect(api.runScheduledCheck).not.toHaveBeenCalled();
  });

  it("notification-delivery / test notification UI / shows native request status and isolated suppression", async () => {
    const api = setupApi();
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    await userEvent.click(
      screen.getByRole("tab", { name: "通知", exact: true }),
    );
    const button = await screen.findByRole("button", { name: "發送測試通知" });
    await userEvent.click(button);
    expect(
      await screen.findByText("隔離測試模式不顯示 Windows 通知。"),
    ).toBeVisible();
    expect(api.sendTestNotification).toHaveBeenCalledOnce();

    vi.mocked(api.sendTestNotification).mockResolvedValueOnce({
      status: "requested",
    });
    await userEvent.click(button);
    expect(
      await screen.findByText(
        "Windows 已接受測試通知顯示要求；若未看到，請檢查系統通知設定。",
      ),
    ).toBeVisible();

    vi.mocked(api.sendTestNotification).mockRejectedValueOnce(
      new Error("系統通知不可用"),
    );
    await userEvent.click(button);
    expect(
      await screen.findByText("測試通知失敗：系統通知不可用"),
    ).toHaveAttribute("role", "alert");
  });

  it("notification-delivery / delayed test notification UI / schedules one minute later and explains the tray test", async () => {
    const api = setupApi();
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    await userEvent.click(
      screen.getByRole("tab", { name: "通知", exact: true }),
    );
    await userEvent.click(
      await screen.findByRole("button", { name: "一分鐘後發送測試通知" }),
    );
    expect(api.scheduleTestNotification).toHaveBeenCalledOnce();
    expect(await screen.findByText(/已排定.*關閉視窗.*系統匣/)).toBeVisible();
  });

  it("monitoring-schedule / settings UI / explains that the interval is measured from the last check, not clock quarter-hours", async () => {
    setupApi();
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    expect(await screen.findByText(/上次檢查.*不是固定整點/)).toBeVisible();
    expect(screen.getByText(/無新公告時不通知/)).toBeVisible();
  });

  it("notification-delivery / test notification UI / returns to settings on a test toast click", async () => {
    setupApi();
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "重大訊息" }));
    (
      window as Window & {
        sendReporterNotificationRoute: (route: {
          type: "test-notification";
        }) => void;
      }
    ).sendReporterNotificationRoute({ type: "test-notification" });

    expect(
      await screen.findByRole("heading", { name: "設定", level: 1 }),
    ).toBeVisible();
    expect(await screen.findByText("已從測試通知返回設定。")).toBeVisible();
  });

  it("shows immediate progress beside the selected job while an actual check is pending", async () => {
    const api = setupApi();
    api.runScheduledCheck = vi.fn(() => new Promise(() => undefined));
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    await userEvent.click(
      await screen.findByRole("button", { name: "立即檢查重大訊息" }),
    );
    const card = screen
      .getByRole("heading", { name: "重大訊息" })
      .closest("section")!;
    expect(within(card).getByRole("status")).toHaveTextContent(
      "正在檢查重大訊息",
    );
    expect(
      within(card).getByRole("button", { name: "檢查中…" }),
    ).toBeDisabled();
  });

  it("reports an invalid code and adds a valid company from the keyboard", async () => {
    const api = setupApi();
    const user = userEvent.setup();
    renderWatchlist();
    await userEvent.click(
      screen.getByRole("button", { name: "新增關注公司", exact: true }),
    );
    const form = screen.getByRole("form", { name: "新增關注公司" });
    const code = within(form).getByRole("textbox", { name: "股票代號" });
    await user.type(code, "bad");
    await user.keyboard("{Enter}");
    expect(await screen.findByRole("alert")).toHaveTextContent(
      "股票代號必須是 4 至 6 位數字",
    );
    expect(api.addWatchlist).not.toHaveBeenCalled();

    await user.clear(code);
    await user.type(code, "2317");
    await user.keyboard("{Enter}");
    await waitFor(() =>
      expect(api.addWatchlist).toHaveBeenCalledWith(
        expect.objectContaining({
          market: "TWSE",
          stockCode: "2317",
        }),
      ),
    );
  });

  it("edits category and notes and preserves the company in the local list when disabling it", async () => {
    const api = setupApi();
    const user = userEvent.setup();
    renderWatchlist();
    await user.click(await screen.findByRole("button", { name: "編輯台積電" }));
    const dialog = screen.getByRole("dialog");
    await user.clear(within(dialog).getByLabelText("分類"));
    await user.type(within(dialog).getByLabelText("分類"), "半導體");
    await user.clear(within(dialog).getByLabelText("備註"));
    await user.type(within(dialog).getByLabelText("備註"), "注意法說會");
    await user.click(
      within(dialog).getByRole("button", { name: "儲存", exact: true }),
    );
    await waitFor(() =>
      expect(api.updateWatchlist).toHaveBeenCalledWith({
        companyId: "company-1",
        category: "半導體",
        notes: "注意法說會",
      }),
    );
    const toggle = screen.getByRole("switch", { name: "監控台積電" });
    await waitFor(() => expect(toggle).toBeEnabled());
    await user.click(toggle);
    await waitFor(() =>
      expect(toggle).toHaveAttribute("aria-checked", "false"),
    );
    expect(screen.getByText("台積電")).toBeVisible();
    expect(api.setWatchlistActive).toHaveBeenCalledWith({
      companyId: "company-1",
      active: false,
    });
  });

  it("previews an old configuration before applying the import", async () => {
    const api = setupApi();
    const user = userEvent.setup();
    renderWatchlist();
    await userEvent.click(
      screen.getByRole("button", { name: "匯入舊版股票清單", exact: true }),
    );
    const config = screen.getByRole("textbox", { name: "舊版設定 JSON" });
    fireEvent.change(config, {
      target: { value: '{"StockNumbers":"2330,2317"}' },
    });
    await user.click(screen.getByRole("button", { name: "預覽匯入" }));
    expect(await screen.findByText("新增 1、略過 1、失敗 1")).toBeVisible();
    expect(api.applyWatchlistImport).not.toHaveBeenCalled();
    await user.click(screen.getByRole("button", { name: "確認匯入" }));
    await waitFor(() =>
      expect(api.applyWatchlistImport).toHaveBeenCalledWith(
        '{"StockNumbers":"2330,2317"}',
      ),
    );
  });

  it("shows every market disclosure while visually highlighting the watched company", async () => {
    const api = setupApi();
    const user = userEvent.setup();
    renderWatchlist();
    await user.click(screen.getByRole("tab", { name: "違約交割" }));
    expect(await screen.findByText("關注半導體")).toBeVisible();
    expect(await screen.findByText("一般公司")).toBeVisible();
    expect(
      screen.getByText("關注公司", { selector: ".watched-badge" }),
    ).toBeVisible();
    expect(api.listDisclosures).toHaveBeenCalledWith(
      expect.objectContaining({ disclosureDate: expect.any(String) }),
    );
  });

  it("shows stale and failed disclosure sources separately from a zero-row list", async () => {
    const api = setupApi();
    vi.mocked(api.listDisclosures).mockResolvedValue([]);
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "違約交割" }));
    expect(
      await screen.findByRole("status", { name: "違約交割來源狀態" }),
    ).toHaveTextContent("資料來源部分降級");
    expect(
      screen.getByRole("status", { name: "違約交割來源狀態" }),
    ).toHaveTextContent(/TWSE.*2026-09-25.*官方尚未更新/);
    expect(
      screen.getByRole("status", { name: "違約交割來源狀態" }),
    ).toHaveTextContent(/原定重試 2026\/09\/26\s*19:00/);
    expect(
      screen.getByRole("status", { name: "違約交割來源狀態" }),
    ).not.toHaveTextContent("預計重試");
    expect(
      screen.getByRole("status", { name: "違約交割來源狀態" }),
    ).not.toHaveTextContent("T19:00:00+08:00");
    expect(
      screen.getByRole("status", { name: "違約交割來源狀態" }),
    ).toHaveTextContent(/TPEX.*來源失敗/);
    expect(
      screen.getByText("來源檢查失敗或官方資料尚未更新，不能判定為零筆。"),
    ).toBeVisible();
    expect(
      screen.getByRole("heading", {
        name: "來源檢查失敗或官方資料尚未更新，不能判定為零筆。",
      }),
    ).toBeVisible();
    expect(
      screen.queryByText("當日無個股達違約資訊揭露標準"),
    ).not.toBeInTheDocument();
    await userEvent.click(screen.getByRole("tab", { name: "總覽" }));
    await waitFor(() =>
      expect(document.querySelector(".overview-summary")).toHaveTextContent(
        /今日違約揭露.*未確認/,
      ),
    );
  });

  it("labels a current-date zero-result as no disclosure rather than an error", async () => {
    const api = setupApi();
    vi.mocked(api.listDisclosures).mockResolvedValue([]);
    vi.mocked(api.getDisclosureMonitorStatus).mockResolvedValue({
      status: "complete",
      finishedAt: "2026-09-26T11:00:00.000Z",
      errorMessage: null,
      markets: {
        TWSE: {
          status: "empty_success",
          dataDate: "2026-09-26",
          recordCount: 0,
          retryAt: null,
        },
      },
    });
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "違約交割" }));
    fireEvent.change(screen.getByLabelText("違約交割資料日期"), {
      target: { value: "2026-09-26" },
    });
    fireEvent.change(screen.getByRole("combobox", { name: "違約交割市場" }), {
      target: { value: "TWSE" },
    });
    expect(
      await screen.findByRole("status", { name: "違約交割來源狀態" }),
    ).toHaveTextContent(/TWSE.*本日無揭露/);
    expect(screen.getByText("當日無個股達違約資訊揭露標準")).toBeVisible();
  });

  it("opens full event details without losing the search and links a correction to its original", async () => {
    const api = setupApi();
    const user = userEvent.setup();
    renderWatchlist();
    await user.click(screen.getByRole("tab", { name: "重大訊息" }));
    const search = screen.getByRole("searchbox", { name: "搜尋重大訊息" });
    await user.type(search, "更正");
    await user.click(
      await screen.findByRole("button", { name: /營收更正公告/ }),
    );
    const expandedRow = screen
      .getByRole("button", { name: /營收更正公告/ })
      .closest(".material-row")!;
    expect(
      within(expandedRow as HTMLElement).getByRole("complementary", {
        name: "重大訊息完整內容",
      }),
    ).toBeVisible();
    expect(await screen.findByText("更正後完整公告內容")).toBeVisible();
    expect(screen.getByRole("link", { name: "開啟來源公告" })).toHaveAttribute(
      "href",
      "https://mops.example/event-2",
    );
    expect(screen.getByText("原公告：營運報告")).toBeVisible();
    expect(screen.getByRole("searchbox", { name: "搜尋重大訊息" })).toHaveValue(
      "更正",
    );
    expect(api.getMaterialEvent).toHaveBeenCalledWith("event-2");
  });

  it("defaults to watched announcements, offers all saved announcements, and expands only the clicked row", async () => {
    const api = setupApi();
    const user = userEvent.setup();
    renderWatchlist();
    await user.click(screen.getByRole("tab", { name: "重大訊息" }));
    await waitFor(() =>
      expect(api.listMaterialEvents).toHaveBeenLastCalledWith(
        expect.objectContaining({ watchedOnly: true }),
      ),
    );
    expect(
      await screen.findByRole("button", { name: /台積電 營收更正公告/ }),
    ).toBeVisible();
    expect(screen.getByText("顯示 2 筆公告")).toBeVisible();
    expect(
      screen.queryByRole("button", { name: /其他公司 其他重大公告/ }),
    ).not.toBeInTheDocument();
    await user.click(screen.getByRole("checkbox", { name: "顯示全部公告" }));
    expect(
      await screen.findByRole("button", { name: /其他公司 其他重大公告/ }),
    ).toBeVisible();
    expect(screen.getByText("顯示 3 筆公告")).toBeVisible();
    await user.click(
      screen.getByRole("button", { name: /台積電 營收更正公告/ }),
    );
    expect(
      within(
        screen
          .getByRole("button", { name: /台積電 營收更正公告/ })
          .closest(".material-row") as HTMLElement,
      ).getByRole("complementary", { name: "重大訊息完整內容" }),
    ).toBeVisible();
    await user.click(
      screen.getByRole("button", { name: /其他公司 其他重大公告/ }),
    );
    expect(
      within(
        screen
          .getByRole("button", { name: /其他公司 其他重大公告/ })
          .closest(".material-row") as HTMLElement,
      ).getByRole("complementary", { name: "重大訊息完整內容" }),
    ).toBeVisible();
    expect(
      within(
        screen
          .getByRole("button", { name: /台積電 營收更正公告/ })
          .closest(".material-row") as HTMLElement,
      ).queryByRole("complementary"),
    ).not.toBeInTheDocument();
  });

  it("material-event-monitoring / soft delete UI / requires confirmation and removes only the chosen row", async () => {
    const api = setupApi();
    const confirm = vi
      .spyOn(window, "confirm")
      .mockReturnValueOnce(false)
      .mockReturnValueOnce(true);
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "重大訊息" }));
    await userEvent.click(
      await screen.findByRole("button", { name: /台積電 營收更正公告/ }),
    );
    const button = await screen.findByRole("button", {
      name: "刪除此筆本機公告",
    });
    await userEvent.click(button);
    expect(confirm).toHaveBeenCalledOnce();
    expect(api.softDeleteMaterialEvent).not.toHaveBeenCalled();
    expect(
      screen.getByRole("button", { name: /台積電 營收更正公告/ }),
    ).toBeVisible();
    await userEvent.click(button);
    expect(api.softDeleteMaterialEvent).toHaveBeenCalledWith("event-2");
    await waitFor(() =>
      expect(
        screen.queryByRole("button", { name: /台積電 營收更正公告/ }),
      ).not.toBeInTheDocument(),
    );
    expect(
      screen.getByRole("button", { name: /台積電 營運報告/ }),
    ).toBeVisible();
    expect(
      screen.getByRole("status", { name: "重大訊息操作結果" }),
    ).toHaveTextContent("已從本機清單移除");
  });

  it("material-event-monitoring / soft delete UI / prevents a duplicate request while deletion is pending", async () => {
    const api = setupApi();
    vi.spyOn(window, "confirm").mockReturnValue(true);
    let complete: (() => void) | undefined;
    vi.mocked(api.softDeleteMaterialEvent).mockImplementation(
      () =>
        new Promise((resolve) => {
          complete = () => resolve(undefined);
        }),
    );
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "重大訊息" }));
    await userEvent.click(
      await screen.findByRole("button", { name: /台積電 營收更正公告/ }),
    );
    const button = await screen.findByRole("button", {
      name: "刪除此筆本機公告",
    });
    await userEvent.click(button);
    expect(button).toBeDisabled();
    await userEvent.click(button);
    expect(api.softDeleteMaterialEvent).toHaveBeenCalledOnce();
    complete?.();
  });

  it("stops loading and reports a detail retrieval error in the clicked row", async () => {
    const api = setupApi();
    vi.mocked(api.getMaterialEvent).mockRejectedValueOnce(
      new Error("公告詳情暫時無法取得"),
    );
    const user = userEvent.setup();
    renderWatchlist();
    await user.click(screen.getByRole("tab", { name: "重大訊息" }));
    await user.click(
      await screen.findByRole("button", { name: /台積電 營收更正公告/ }),
    );
    expect(await screen.findByText("公告詳情暫時無法取得")).toHaveAttribute(
      "role",
      "alert",
    );
    expect(screen.queryByText("載入完整公告…")).not.toBeInTheDocument();
    expect(
      screen.getByRole("button", { name: /台積電 營收更正公告/ }),
    ).toHaveAttribute("aria-expanded", "false");
  });

  it("shows a source failure separately from an empty material-event result", async () => {
    const api = setupApi();
    renderWatchlist();
    vi.mocked(api.listMaterialEvents).mockResolvedValue([]);
    await userEvent.click(screen.getByRole("tab", { name: "重大訊息" }));
    await userEvent.click(
      await screen.findByRole("button", { name: "展開詳情" }),
    );
    expect(
      await screen.findByRole("status", { name: "重大訊息來源狀態" }),
    ).toHaveTextContent("資料來源部分降級");
    expect(
      screen.getByRole("status", { name: "重大訊息來源狀態" }),
    ).toHaveTextContent("MOPS 暫時無法連線");
    expect(
      screen.getByText(
        "目前沒有符合條件的關注公司重大訊息；可勾選「顯示全部公告」。",
      ),
    ).toBeVisible();
    expect(
      screen.getByRole("status", { name: "重大訊息來源狀態" }),
    ).toHaveTextContent("上市結構化每日對帳：資料完整");
  });

  it("opens a material event detail with keyboard activation", async () => {
    const api = setupApi();
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "重大訊息" }));
    const eventButton = await screen.findByRole("button", {
      name: /台積電 營收更正公告/,
    });
    eventButton.focus();
    await userEvent.keyboard("{Enter}");
    expect(
      await screen.findByRole("complementary", { name: "重大訊息完整內容" }),
    ).toHaveTextContent("更正後完整公告內容");
    expect(api.getMaterialEvent).toHaveBeenCalledWith("event-2");
  });

  it("routes a source-failure toast to the material source status without requesting an event", async () => {
    const api = setupApi();
    vi.mocked(api.listMaterialEvents).mockResolvedValue([]);
    renderWatchlist();
    (
      window as Window & {
        sendReporterNotificationRoute: (route: {
          type: "source-status";
        }) => void;
      }
    ).sendReporterNotificationRoute({ type: "source-status" });
    expect(
      await screen.findByRole("status", { name: "重大訊息來源狀態" }),
    ).toHaveTextContent("資料來源部分降級");
    expect(api.getMaterialEvent).not.toHaveBeenCalled();
  });

  it("routes a single-event toast directly to the full event details", async () => {
    const api = setupApi();
    renderWatchlist();
    (
      window as Window & {
        sendReporterNotificationRoute: (route: {
          type: "event-detail";
          eventId: string;
        }) => void;
      }
    ).sendReporterNotificationRoute({
      type: "event-detail",
      eventId: "event-2",
    });
    expect(
      await screen.findByRole("complementary", { name: "重大訊息完整內容" }),
    ).toHaveTextContent("更正後完整公告內容");
    expect(api.getMaterialEvent).toHaveBeenCalledWith("event-2");
  });

  it("opens a filtered material list from a notification batch route", async () => {
    const api = setupApi();
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "重大訊息" }));
    await screen.findByRole("button", { name: /台積電 營收更正公告/ });
    (
      window as Window & {
        sendReporterNotificationRoute: (route: {
          type: "event-list";
          eventIds: string[];
        }) => void;
      }
    ).sendReporterNotificationRoute({
      type: "event-list",
      eventIds: ["event-2"],
    });
    expect(
      await screen.findByRole("button", { name: /台積電 營收更正公告/ }),
    ).toBeVisible();
    await waitFor(() =>
      expect(api.listMaterialEvents).toHaveBeenLastCalledWith(
        expect.objectContaining({ eventIds: ["event-2"] }),
      ),
    );
    expect(
      screen.queryByRole("button", { name: /台積電 營運報告/ }),
    ).not.toBeInTheDocument();
  });
});

describe("desktop-app-lifecycle / login startup UI", () => {
  afterEach(() => cleanup());

  it("persists a login-startup toggle immediately so a checked control is not mistaken for a saved setting", async () => {
    const api = setupApi();
    const user = userEvent.setup();
    renderWatchlist();
    await user.click(screen.getByRole("tab", { name: "設定" }));
    await userEvent.click(
      screen.getByRole("tab", { name: "啟動與系統匣", exact: true }),
    );
    await user.click(
      await screen.findByRole("switch", { name: "登入 Windows 時啟動" }),
    );
    await waitFor(() =>
      expect(api.saveLoginStartupSettings).toHaveBeenCalledWith({
        enabled: true,
        startHidden: false,
      }),
    );
    expect(await screen.findByText("登入啟動設定已儲存。")).toBeVisible();
  });

  it("loads, enables and saves login launch and hidden-start preferences", async () => {
    const api = setupApi();
    const user = userEvent.setup();
    renderWatchlist();
    await user.click(screen.getByRole("tab", { name: "設定" }));
    await userEvent.click(
      screen.getByRole("tab", { name: "啟動與系統匣", exact: true }),
    );
    expect(await screen.findByLabelText("登入 Windows 時啟動")).toBeVisible();
    await user.click(screen.getByLabelText("登入 Windows 時啟動"));
    await waitFor(() =>
      expect(api.saveLoginStartupSettings).toHaveBeenCalledWith({
        enabled: true,
        startHidden: false,
      }),
    );
    await waitFor(() =>
      expect(screen.getByLabelText("啟動後縮到系統匣")).toBeEnabled(),
    );
    await user.click(screen.getByLabelText("啟動後縮到系統匣"));
    await waitFor(() =>
      expect(api.saveLoginStartupSettings).toHaveBeenCalledWith({
        enabled: true,
        startHidden: true,
      }),
    );
  });

  it("keeps the login-startup toggle off and reports an OS save failure", async () => {
    const api = setupApi();
    api.saveLoginStartupSettings.mockRejectedValueOnce(
      new Error("Windows 登入啟動未生效"),
    );
    renderWatchlist();
    await userEvent.click(screen.getByRole("tab", { name: "設定" }));
    await userEvent.click(
      screen.getByRole("tab", { name: "啟動與系統匣", exact: true }),
    );
    const toggle = await screen.findByRole("switch", {
      name: "登入 Windows 時啟動",
    });
    await userEvent.click(toggle);
    expect(
      await screen.findByText("登入啟動設定失敗：Windows 登入啟動未生效"),
    ).toBeVisible();
    expect(toggle).not.toBeChecked();
  });
});
