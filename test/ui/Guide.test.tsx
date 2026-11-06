import { useState } from "react";
import {
  act,
  cleanup,
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { Config } from "driver.js";
import {
  GuideProvider,
  GuideWelcome,
  GuideLauncher,
} from "../../src/renderer/Guide";
import { PageHeader } from "../../src/renderer/ui";
import {
  GUIDE_STORAGE_KEY,
  pageGuides,
  type GuidePage,
} from "../../src/renderer/onboarding";

const mocked = vi.hoisted(() => ({
  configs: [] as Config[],
  destroy: vi.fn(),
  drive: vi.fn(),
}));
vi.mock("driver.js", () => ({
  driver: (config: Config) => {
    mocked.configs.push(config);
    return { drive: mocked.drive, destroy: mocked.destroy };
  },
}));
const business = vi.fn();
function Harness() {
  const [page, setPage] = useState<GuidePage>("overview");
  return (
    <GuideProvider page={page} onNavigate={setPage}>
      <GuideLauncher />
      <GuideWelcome />
      <button onClick={() => setPage("overview")}>通知換頁</button>
      <PageHeader title={pageGuides[page].title} subtitle="測試頁面">
        <button onClick={business}>業務操作</button>
      </PageHeader>
    </GuideProvider>
  );
}
const config = () => mocked.configs.at(-1)!;
async function next() {
  await act(async () => {
    config().onNextClick?.(undefined as never, {} as never, {} as never);
  });
}
async function previous() {
  await act(async () => {
    config().onPrevClick?.(undefined as never, {} as never, {} as never);
  });
}
async function close() {
  await act(async () => {
    config().onDestroyStarted?.(undefined as never, {} as never, {} as never);
  });
}

beforeEach(() => {
  localStorage.clear();
  mocked.configs.length = 0;
  vi.clearAllMocks();
});
afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});
describe("guided onboarding", () => {
  it("dismisses the welcome card, persists skip and can reopen the guide", () => {
    const view = render(<Harness />);
    expect(
      screen.getByRole("region", { name: "歡迎使用小幫手" }),
    ).toBeVisible();
    fireEvent.click(screen.getByRole("button", { name: "稍後" }));
    expect(screen.queryByRole("region", { name: "歡迎使用小幫手" })).toBeNull();
    view.unmount();
    render(<Harness />);
    expect(screen.queryByRole("region", { name: "歡迎使用小幫手" })).toBeNull();
    fireEvent.click(screen.getByRole("button", { name: "使用指南" }));
    expect(screen.getByRole("dialog", { name: "使用指南" })).toBeVisible();
  });
  it("explains all seven pages and exposes each page tour", () => {
    render(<Harness />);
    for (const guide of Object.values(pageGuides)) {
      fireEvent.click(screen.getByRole("button", { name: "使用指南" }));
      fireEvent.click(
        within(screen.getByRole("dialog", { name: "使用指南" })).getByRole(
          "button",
          { name: guide.title, exact: true },
        ),
      );
      const dialog = screen.getByRole("dialog", {
        name: `${guide.title}使用說明`,
      });
      expect(within(dialog).getByText(guide.summary)).toBeVisible();
      for (const step of guide.steps)
        expect(within(dialog).getByText(step.description)).toBeVisible();
      expect(
        within(dialog).getByRole("button", { name: "帶我操作" }),
      ).toBeVisible();
      fireEvent.click(
        within(dialog).getByRole("button", { name: "關閉對話框" }),
      );
    }
  });
  it("navigates five quick steps both ways, completes, and stays completed after remount", async () => {
    const view = render(<Harness />);
    fireEvent.click(screen.getByRole("button", { name: "開始導覽" }));
    await waitFor(() => expect(mocked.drive).toHaveBeenCalled());
    expect(screen.getByRole("heading", { level: 1 })).toHaveTextContent(
      "關注公司",
    );
    expect(config().progressText).toBe("第 1 / 5 步");
    await next();
    expect(screen.getByRole("heading", { level: 1 })).toHaveTextContent(
      "重大訊息",
    );
    await previous();
    expect(config().progressText).toBe("第 1 / 5 步");
    for (let i = 0; i < 5; i++) await next();
    expect(JSON.parse(localStorage.getItem(GUIDE_STORAGE_KEY)!)).toMatchObject({
      status: "complete",
    });
    expect(business).not.toHaveBeenCalled();
    view.unmount();
    render(<Harness />);
    expect(screen.queryByRole("region", { name: "歡迎使用小幫手" })).toBeNull();
    fireEvent.click(screen.getByRole("button", { name: "使用指南" }));
    fireEvent.click(screen.getByRole("button", { name: "重看快速導覽" }));
    await waitFor(() => expect(config().progressText).toBe("第 1 / 5 步"));
  });
  it("opens page help and survives absent targets, then cleans up and restores focus", async () => {
    render(<Harness />);
    fireEvent.click(screen.getByRole("button", { name: "使用說明" }));
    fireEvent.click(screen.getByRole("button", { name: "帶我操作" }));
    await waitFor(() => expect(config().progressText).toBe("第 1 / 3 步"));
    await next();
    expect(config().steps?.[1].element).toBeUndefined();
    await close();
    expect(mocked.destroy).toHaveBeenCalled();
    await waitFor(() =>
      expect(screen.getByRole("button", { name: "使用說明" })).toHaveFocus(),
    );
    expect(business).not.toHaveBeenCalled();
  });
  it("reports when preferences cannot be saved", () => {
    vi.spyOn(Storage.prototype, "setItem").mockImplementation(() => {
      throw Error("unavailable");
    });
    render(<Harness />);
    fireEvent.click(screen.getByRole("button", { name: "稍後" }));
    expect(screen.getByRole("status")).toHaveTextContent("導覽偏好無法保存");
  });
  it("cancels on an external page change instead of navigating back", async () => {
    const view = render(<Harness />);
    fireEvent.click(screen.getByRole("button", { name: "開始導覽" }));
    await waitFor(() => expect(mocked.drive).toHaveBeenCalled());
    fireEvent.click(screen.getByRole("button", { name: "通知換頁" }));
    expect(screen.getByRole("heading", { level: 1 })).toHaveTextContent("總覽");
    expect(mocked.destroy).toHaveBeenCalled();
    const calls = mocked.drive.mock.calls.length;
    view.unmount();
    expect(mocked.drive).toHaveBeenCalledTimes(calls);
  });
});
