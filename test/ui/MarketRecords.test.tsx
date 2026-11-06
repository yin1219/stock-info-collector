import {
  render,
  screen,
  fireEvent,
  waitFor,
  cleanup,
  act,
} from "@testing-library/react";
import { describe, it, expect, vi, afterEach } from "vitest";
import { MarketRecords } from "../../src/renderer/MarketRecords";
describe("market-records UI", () => {
  afterEach(cleanup);
  it("refreshes visible background data, preserves the chosen date, and stops reading on unmount", async () => {
    vi.useFakeTimers();
    const visible = vi
      .spyOn(document, "visibilityState", "get")
      .mockReturnValue("visible");
    const base = {
      today: "2026-11-06",
      todayConfirmed: false,
      coverage: { start: "2020-01-02", end: "2026-11-05" },
      indices: { rise: [], fall: [] },
      flowGroups: [],
      news: [],
      manualRequired: false,
      error: null,
      indexUpdate: {
        automatic: true,
        latestOfficialDate: "2026-11-05",
        checkedAt: null,
        knownMissingDates: [],
      },
    };
    const updated = {
      ...base,
      todayConfirmed: true,
      indexUpdate: { ...base.indexUpdate, latestOfficialDate: "2026-11-06" },
    };
    const api = {
      getMarketRecords: vi
        .fn()
        .mockResolvedValueOnce(base)
        .mockResolvedValue(updated),
      refreshMarketRecords: vi.fn(),
      importMarketRecords: vi.fn(),
    };
    try {
      const view = render(<MarketRecords api={api} />);
      await act(async () => {
        await Promise.resolve();
      });
      expect(
        screen.getByRole("region", { name: "指數資料更新" }),
      ).toHaveTextContent("2026-11-05");
      fireEvent.change(screen.getByLabelText("盤面資料日期"), {
        target: { value: "2026-11-04" },
      });
      await act(async () => {
        await vi.advanceTimersByTimeAsync(60_000);
      });
      expect(api.getMarketRecords).toHaveBeenCalledTimes(2);
      expect(
        screen.getByRole("region", { name: "指數資料更新" }),
      ).toHaveTextContent("2026-11-06");
      expect(screen.getByLabelText("盤面資料日期")).toHaveValue("2026-11-04");
      visible.mockReturnValue("hidden");
      await act(async () => {
        await vi.advanceTimersByTimeAsync(60_000);
      });
      expect(api.getMarketRecords).toHaveBeenCalledTimes(2);
      visible.mockReturnValue("visible");
      await act(async () => {
        window.dispatchEvent(new Event("focus"));
        await Promise.resolve();
      });
      expect(api.getMarketRecords).toHaveBeenCalledTimes(3);
      view.unmount();
      await vi.advanceTimersByTimeAsync(60_000);
      window.dispatchEvent(new Event("focus"));
      expect(api.getMarketRecords).toHaveBeenCalledTimes(3);
    } finally {
      visible.mockRestore();
      vi.useRealTimers();
    }
  });

  it("keeps the newer manual result when a delayed background snapshot arrives last", async () => {
    vi.useFakeTimers();
    const visible = vi
      .spyOn(document, "visibilityState", "get")
      .mockReturnValue("visible");
    const base = {
      today: "2026-11-06",
      todayConfirmed: false,
      coverage: { start: "2020-01-02", end: "2026-11-05" },
      indices: { rise: [], fall: [] },
      flowGroups: [],
      news: [],
      manualRequired: false,
      error: null,
      indexUpdate: {
        automatic: true,
        latestOfficialDate: "2026-11-05",
        checkedAt: null,
        knownMissingDates: [],
      },
    };
    const updated = {
      ...base,
      todayConfirmed: true,
      indices: {
        rise: [{ date: "2026-11-06", amount: 1234, rank: 1 }],
        fall: [],
      },
      indexUpdate: { ...base.indexUpdate, latestOfficialDate: "2026-11-06" },
    };
    let release!: (value: unknown) => void;
    const pending = new Promise((resolve) => {
      release = resolve;
    });
    const api = {
      getMarketRecords: vi
        .fn()
        .mockResolvedValueOnce(base)
        .mockReturnValueOnce(pending),
      refreshMarketRecords: vi.fn().mockResolvedValue(updated),
      importMarketRecords: vi.fn(),
    };
    try {
      const view = render(<MarketRecords api={api} />);
      await act(async () => {
        await Promise.resolve();
      });
      await act(async () => {
        await vi.advanceTimersByTimeAsync(60_000);
      });
      expect(api.getMarketRecords).toHaveBeenCalledTimes(2);
      await act(async () => {
        fireEvent.click(screen.getByRole("button", { name: "更新官方指數" }));
        await Promise.resolve();
      });
      expect(
        screen.getByRole("region", { name: "指數資料更新" }),
      ).toHaveTextContent("2026-11-06");
      await act(async () => {
        release(base);
        await pending;
      });
      expect(
        screen.getByRole("region", { name: "指數資料更新" }),
      ).toHaveTextContent("2026-11-06");
      expect(screen.getAllByText(/\+1,234\.00/)).not.toHaveLength(0);
      view.unmount();
    } finally {
      visible.mockRestore();
      vi.useRealTimers();
    }
  });

  it("shows daily OpenAPI status and the CMoney historical-gap import entry", async () => {
    const snapshot = {
      today: "2026-11-06",
      todayConfirmed: false,
      coverage: { start: "2020-01-02", end: "2026-11-05" },
      indices: { rise: [], fall: [] },
      flowGroups: [],
      news: [],
      manualRequired: false,
      error: null,
      indexUpdate: {
        automatic: true,
        latestOfficialDate: "2026-11-05",
        checkedAt: "2026-11-06T08:00:00Z",
        knownMissingDates: ["2026-11-03"],
      },
    };
    const api = {
      getMarketRecords: vi.fn().mockResolvedValue(snapshot),
      refreshMarketRecords: vi.fn().mockResolvedValue(snapshot),
      importMarketRecords: vi.fn().mockResolvedValue({ status: "cancelled" }),
    };
    render(<MarketRecords api={api} />);
    const status = await screen.findByRole("region", { name: "指數資料更新" });
    expect(status).toHaveTextContent("15:30");
    expect(status).toHaveTextContent("2026-11-05");
    expect(status).toHaveTextContent("2026-11-03");
    fireEvent.click(screen.getByRole("button", { name: "匯入 CMoney 補齊" }));
    expect(screen.getByRole("dialog", { name: "盤面資料管理" })).toBeVisible();
    fireEvent.click(screen.getByRole("button", { name: "匯入加權指數 Excel" }));
    await waitFor(() =>
      expect(api.importMarketRecords).toHaveBeenCalledWith("index"),
    );
  });

  it("renders actual index and investor data in separate tabs and opens a shared import dialog", async () => {
    const snapshot = {
      today: "2026-11-05",
      todayConfirmed: true,
      coverage: { start: "2020-01-02", end: "2026-11-05" },
      indices: {
        rise: [{ date: "2026-11-05", amount: 1000, rank: 2 }],
        fall: [],
      },
      flowGroups: [
        {
          metric: "trust-buy",
          label: "投信連續買超",
          source: "CMoney",
          rows: [
            {
              start: "2026-11-03",
              end: "2026-11-05",
              days: 3,
              amount: 1200000000,
              rank: 1,
            },
          ],
        },
      ],
      news: [],
      error: null,
      manualRequired: false,
      sourceStatus: "已保存日資料",
      foreignWarning: null,
      revisions: 0,
    };
    const api = {
      getMarketRecords: vi.fn().mockResolvedValue(snapshot),
      refreshMarketRecords: vi.fn().mockResolvedValue(snapshot),
      importMarketRecords: vi.fn().mockResolvedValue({ status: "cancelled" }),
    };
    render(<MarketRecords api={api} />);
    expect(
      await screen.findByRole("tab", { name: "加權指數", exact: true }),
    ).toHaveAttribute("aria-selected", "true");
    expect(await screen.findAllByText(/1,000/)).not.toHaveLength(0);
    fireEvent.click(screen.getByRole("tab", { name: "三大法人", exact: true }));
    fireEvent.click(screen.getByRole("tab", { name: "投信", exact: true }));
    expect(
      screen.getByRole("heading", { name: /連續買超前三名/ }),
    ).toBeVisible();
    expect(screen.getByText("12.00")).toBeVisible();
    fireEvent.click(
      screen.getByRole("button", { name: "資料管理", exact: true }),
    );
    expect(screen.getByRole("dialog", { name: "盤面資料管理" })).toBeVisible();
    fireEvent.click(screen.getByRole("button", { name: "匯入投信 Excel" }));
    await waitFor(() =>
      expect(api.importMarketRecords).toHaveBeenCalledWith("trust"),
    );
  });
  it("shows missing data, manual CSV fallback and refresh feedback", async () => {
    const snapshot = {
      today: "2026-11-05",
      manualRequired: true,
      error: null,
      flowGroups: [],
      indices: { rise: [], fall: [] },
      news: [],
      coverage: { start: null, end: null },
      todayConfirmed: false,
    };
    const api = {
      getMarketRecords: vi.fn().mockResolvedValue(snapshot),
      refreshMarketRecords: vi.fn().mockResolvedValue(snapshot),
      importMarketRecords: vi.fn().mockResolvedValue({ status: "cancelled" }),
    };
    render(<MarketRecords api={api} />);
    expect(
      await screen.findByText("以 2020 年以來的日資料比對"),
    ).toBeInTheDocument();
    expect(
      screen.queryByRole("heading", { name: "新聞交叉查核" }),
    ).not.toBeInTheDocument();
    expect(screen.queryByText("未找到報導佐證")).not.toBeInTheDocument();
    expect(screen.getByText(/尚未匯入歷史日資料/)).toBeInTheDocument();
    fireEvent.click(
      screen.getByRole("button", { name: "資料管理", exact: true }),
    );
    expect(screen.getByText(/不必每次重匯完整歷史/)).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "完成", exact: true }));
    fireEvent.click(screen.getByRole("button", { name: "更新官方指數" }));
    await waitFor(() =>
      expect(api.refreshMarketRecords).toHaveBeenCalledOnce(),
    );
    fireEvent.click(
      screen.getByRole("button", { name: "資料管理", exact: true }),
    );
    fireEvent.click(screen.getByRole("button", { name: "匯入 TWSE 當日 CSV" }));
    await waitFor(() =>
      expect(api.importMarketRecords).toHaveBeenCalledWith("csv"),
    );
  });
});
