import {
  act,
  cleanup,
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  SetupChecklist,
  type SetupApi,
} from "../../src/renderer/SetupChecklist";

afterEach(cleanup);
function fixture() {
  return {
    listWatchlist: vi.fn(async () => [{ active: false }]),
    getScheduleSettings: vi.fn(async () => ({
      monitoring: {
        enabled: true,
        start: "08:00",
        end: "18:00",
        intervalMinutes: 15 as const,
      },
      disclosure: { enabled: true, runAt: "18:30" },
      notifyEmptyDefaultDisclosures: true,
    })),
    getLoginStartupSettings: vi.fn(async () => ({
      enabled: true,
      startHidden: true,
      blockedByWindows: true,
    })),
    getGoogleCalendarStatus: vi.fn(async () => ({
      status: "disconnected" as const,
    })),
    getMarketRecords: vi.fn(async () => ({
      coverage: { start: "2026-11-02", end: "2026-11-05" },
    })),
  } satisfies SetupApi;
}
describe("first-use setup checklist", () => {
  it("reads actual states and distinguishes optional features from completed setup", async () => {
    const api = fixture();
    render(<SetupChecklist api={api} onNavigate={vi.fn()} onClose={vi.fn()} />);
    expect(await screen.findByText("尚無啟用中的公司")).toBeVisible();
    expect(screen.getByText(/Windows 已阻擋登入啟動/)).toBeVisible();
    expect(screen.getByText(/08:00.*18:00.*15 分鐘/)).toBeVisible();
    expect(screen.getByText("尚未連線")).toBeVisible();
    expect(screen.getByText(/2026-11-02.*2026-11-05/)).toBeVisible();
    expect(screen.getByText(/已有日資料不代表歷史完整/)).toBeVisible();
    expect(screen.getByText("需手動測試")).toBeVisible();
    expect(
      screen.getByRole("heading", { name: "依工作需要選擇" }),
    ).toBeVisible();
    expect(screen.getByRole("button", { name: "先使用小幫手" })).toBeEnabled();
  });
  it("opens the exact settings tab and closes the checklist", async () => {
    const navigate = vi.fn(),
      close = vi.fn();
    const api = fixture();
    render(<SetupChecklist api={api} onNavigate={navigate} onClose={close} />);
    await screen.findByText("需手動測試");
    for (const [button, page, tab] of [
      ["設定檢查排程", "settings", "監控排程"],
      ["設定背景啟動", "settings", "啟動與系統匣"],
      ["前往測試通知", "settings", "通知"],
      ["管理關注公司", "watchlist", undefined],
      ["設定 Google Calendar", "conferences", undefined],
      ["匯入歷史日資料", "market-records", undefined],
    ] as const) {
      fireEvent.click(screen.getByRole("button", { name: button }));
      expect(navigate).toHaveBeenLastCalledWith(page, tab);
    }
    expect(close).toHaveBeenCalledTimes(6);
    for (const method of Object.values(api))
      expect(method).toHaveBeenCalledTimes(1);
  });
  it("keeps other states when a read fails and refreshes after configuration changes", async () => {
    const api = fixture();
    api.getGoogleCalendarStatus.mockRejectedValueOnce(Error("network"));
    render(<SetupChecklist api={api} onNavigate={vi.fn()} onClose={vi.fn()} />);
    const google = await screen.findByRole("region", {
      name: "Google Calendar",
    });
    await waitFor(() =>
      expect(within(google).getByText("無法讀取目前狀態")).toBeVisible(),
    );
    expect(screen.getByText("尚無啟用中的公司")).toBeVisible();
    api.listWatchlist.mockResolvedValue([{ active: true }]);
    fireEvent.click(screen.getByRole("button", { name: "重新檢查" }));
    expect(await screen.findByText("已啟用 1 家公司")).toBeVisible();
    expect(within(google).getByText("尚未連線")).toBeVisible();
  });
  it("discards an older pending read after refresh", async () => {
    const api = fixture();
    let resolve!: (value: unknown) => void;
    api.listWatchlist.mockImplementationOnce(
      () =>
        new Promise((r) => {
          resolve = r;
        }) as never,
    );
    render(<SetupChecklist api={api} onNavigate={vi.fn()} onClose={vi.fn()} />);
    await waitFor(() => expect(api.listWatchlist).toHaveBeenCalledTimes(1));
    api.listWatchlist.mockResolvedValue([{ active: true }]);
    fireEvent.click(screen.getByRole("button", { name: "重新檢查" }));
    expect(await screen.findByText("已啟用 1 家公司")).toBeVisible();
    await act(async () => {
      resolve([]);
    });
    expect(screen.getByText("已啟用 1 家公司")).toBeVisible();
  });
});
