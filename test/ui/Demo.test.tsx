import { describe, it, expect, afterEach, vi } from "vitest";
import {
  render,
  screen,
  fireEvent,
  within,
  cleanup,
  act,
} from "@testing-library/react";
import { DemoApp } from "../../src/demo/DemoApp";
afterEach(cleanup);
describe("reporter UI demo", () => {
  it("opens and closes an announcement from the time, company and row padding", () => {
    render(<DemoApp />);
    fireEvent.click(
      screen.getByRole("button", { name: "重大訊息", exact: true }),
    );
    fireEvent.click(screen.getByText("11/05 14:32", { exact: true }));
    expect(screen.getByText("公告內容")).toBeInTheDocument();
    fireEvent.click(screen.getByText("台積電", { exact: true }));
    expect(screen.queryByText("公告內容")).not.toBeInTheDocument();
    fireEvent.click(
      screen.getByRole("button", { name: "董事會決議股利分配", exact: true }),
    );
    expect(screen.getByText("公告內容")).toBeInTheDocument();
  });
  it("explains CMoney export layouts without dated filenames or full-history reimports", () => {
    render(<DemoApp />);
    fireEvent.click(
      screen.getByRole("button", { name: "盤面紀錄", exact: true }),
    );
    for (const tab of ["加權指數", "三大法人"]) {
      fireEvent.click(screen.getByRole("tab", { name: tab, exact: true }));
      fireEvent.click(
        screen.getByRole("button", { name: "資料管理", exact: true }),
      );
      const dialog = screen.getByRole("dialog");
      expect(within(dialog).queryByText(/2020年/)).not.toBeInTheDocument();
      expect(
        within(dialog).getByText(/CMoney 預設表格匯出/),
      ).toBeInTheDocument();
      expect(
        within(dialog).getByText(/不必每次重匯完整歷史/),
      ).toBeInTheDocument();
      expect(
        within(dialog).getByText(/不同分頁共用同一份資料/),
      ).toBeInTheDocument();
      expect(within(dialog).getByText(/逐份選擇類型匯入/)).toBeInTheDocument();
      expect(within(dialog).getAllByText(/上市買賣超金額/)).toHaveLength(3);
      for (const type of ["加權指數", "外資", "投信", "自營商"])
        expect(
          within(dialog).getByLabelText(`選擇${type} Excel`),
        ).toHaveAttribute("accept", ".xlsx");
      fireEvent.click(
        within(dialog).getByRole("button", { name: "完成", exact: true }),
      );
    }
  });
  it("keeps an opened unread announcement visible until it is closed", () => {
    render(<DemoApp />);
    fireEvent.click(
      screen.getByRole("button", { name: "重大訊息", exact: true }),
    );
    fireEvent.click(screen.getByLabelText("只看未讀"));
    fireEvent.click(
      screen.getByRole("button", { name: "董事會決議股利分配", exact: true }),
    );
    expect(screen.getByText("公告內容")).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "關閉公告內容" }));
    expect(
      screen.queryByRole("button", { name: "董事會決議股利分配", exact: true }),
    ).not.toBeInTheDocument();
  });
  it("uses current company monitoring state to filter announcements", () => {
    render(<DemoApp />);
    fireEvent.click(
      screen.getByRole("button", { name: "關注公司", exact: true }),
    );
    fireEvent.click(screen.getByRole("switch", { name: "監控台積電" }));
    fireEvent.click(
      screen.getByRole("button", { name: "重大訊息", exact: true }),
    );
    expect(
      screen.queryByRole("button", { name: "董事會決議股利分配", exact: true }),
    ).not.toBeInTheDocument();
    fireEvent.click(screen.getByLabelText("僅顯示關注公司"));
    expect(
      screen.getByRole("button", { name: "董事會決議股利分配", exact: true }),
    ).toBeInTheDocument();
  });
  it("keeps the delayed demo notification scheduled while changing pages", async () => {
    vi.useFakeTimers();
    try {
      render(<DemoApp />);
      fireEvent.click(
        screen.getByRole("button", { name: "設定", exact: true }),
      );
      fireEvent.click(screen.getByRole("tab", { name: "通知", exact: true }));
      fireEvent.click(
        screen.getByRole("button", {
          name: "一分鐘後發送測試通知",
          exact: true,
        }),
      );
      fireEvent.click(
        screen.getByRole("button", { name: "總覽", exact: true }),
      );
      await act(() => vi.advanceTimersByTimeAsync(60000));
      expect(screen.getByRole("status")).toHaveTextContent(
        "一分鐘測試：示意通知已到期",
      );
    } finally {
      cleanup();
      vi.useRealTimers();
    }
  });
  it("provides seven pages and keeps news verification out of the user interface", () => {
    render(<DemoApp />);
    for (const name of [
      "總覽",
      "關注公司",
      "重大訊息",
      "違約交割",
      "法說會",
      "盤面紀錄",
      "設定",
    ]) {
      fireEvent.click(screen.getByRole("button", { name, exact: true }));
      expect(
        screen.getByRole("heading", { level: 1, name, exact: true }),
      ).toBeInTheDocument();
    }
    fireEvent.click(
      screen.getByRole("button", { name: "盤面紀錄", exact: true }),
    );
    expect(screen.queryByText("新聞交叉查核")).not.toBeInTheDocument();
    fireEvent.click(screen.getByRole("tab", { name: "三大法人" }));
    fireEvent.click(screen.getByRole("tab", { name: "投信", exact: true }));
    expect(screen.getAllByText("3,044.45")).toHaveLength(2);
  });
  it("creates and edits local watchlist companies and confirms removal", () => {
    render(<DemoApp />);
    fireEvent.click(
      screen.getByRole("button", { name: "關注公司", exact: true }),
    );
    fireEvent.click(screen.getByRole("button", { name: "新增關注公司" }));
    const dialog = screen.getByRole("dialog");
    fireEvent.change(within(dialog).getByLabelText("股票代號"), {
      target: { value: "9999" },
    });
    fireEvent.change(within(dialog).getByLabelText("公司名稱"), {
      target: { value: "示例公司" },
    });
    fireEvent.click(within(dialog).getByRole("button", { name: "新增" }));
    expect(screen.getByText("示例公司")).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "更多示例公司操作" }));
    fireEvent.click(screen.getByRole("button", { name: "移除示例公司" }));
    expect(screen.getByRole("dialog")).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "確認移除" }));
    expect(screen.queryByText("示例公司")).not.toBeInTheDocument();
  });
  it("expands announcements in place and makes settings controls respond", () => {
    render(<DemoApp />);
    fireEvent.click(
      screen.getByRole("button", { name: "重大訊息", exact: true }),
    );
    fireEvent.click(screen.getByRole("button", { name: "董事會決議股利分配" }));
    expect(screen.getByText("公告內容")).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "關閉公告內容" }));
    expect(screen.queryByText("公告內容")).not.toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "設定", exact: true }));
    fireEvent.change(screen.getByLabelText("檢查頻率"), {
      target: { value: "30" },
    });
    fireEvent.click(screen.getByRole("button", { name: "儲存變更" }));
    expect(screen.getByRole("status")).toHaveTextContent("已儲存示意設定");
    fireEvent.click(screen.getByRole("tab", { name: "通知", exact: true }));
    fireEvent.click(screen.getByRole("button", { name: "發送測試通知" }));
    expect(screen.getByRole("status")).toHaveTextContent("示意通知");
  });
});
