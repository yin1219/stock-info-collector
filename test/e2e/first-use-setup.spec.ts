import { mkdtemp, mkdir, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import Database from "better-sqlite3";
import { _electron as electron, expect, test } from "@playwright/test";

test("first-use setup / actual saved states, optional features and exact settings routes", async () => {
  test.setTimeout(120_000);
  const userData = await mkdtemp(path.join(os.tmpdir(), "reporter-setup-"));
  const app = await electron.launch({
    args: [process.cwd()],
    cwd: process.cwd(),
    env: {
      ...process.env,
      REPORTER_USER_DATA_DIR: userData,
      REPORTER_TEST_TRAY: "0",
    },
  });
  try {
    const page = await app.firstWindow();
    await page.setViewportSize({ width: 1487, height: 1058 });
    await app.evaluate(({ ipcMain }) => {
      const writes: string[] = [];
      (globalThis as unknown as { setupWrites: string[] }).setupWrites = writes;
      for (const channel of [
        "calendar:connect",
        "calendar:sync",
        "market-records:refresh",
        "market-records:import",
        "desktop:save-login-startup",
        "schedule:save-settings",
        "notification:send-test",
        "notification:schedule-test",
        "watchlist:add",
        "watchlist:import-apply",
        "conference-task:ensure",
        "conference-task:switch-disabled-legacy",
      ]) {
        ipcMain.removeHandler(channel);
        ipcMain.handle(channel, () => {
          writes.push(channel);
          throw Error("unexpected setup side effect");
        });
      }
    });
    await page.getByRole("button", { name: "開始設定", exact: true }).click();
    const dialog = page.getByRole("dialog", {
      name: "首次使用設定",
      exact: true,
    });
    await expect(dialog).toContainText("尚無啟用中的公司");
    await expect(dialog).toContainText("尚未匯入日資料");
    await expect(dialog).toContainText("需手動測試");
    const output = path.join(process.cwd(), "artifacts/onboarding");
    await mkdir(output, { recursive: true });
    await page.screenshot({ path: path.join(output, "setup-desktop.png") });
    const db = new Database(path.join(userData, "reporter.sqlite3"));
    try {
      db.prepare(
        "INSERT INTO companies(id,market,stock_code,name,updated_at) VALUES(?,?,?,?,?)",
      ).run("test-company", "TWSE", "2330", "測試公司", "2026-11-06");
      db.prepare(
        "INSERT INTO watchlist_entries(id,company_id,active,category,notes,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",
      ).run(
        "test-watch",
        "test-company",
        1,
        "",
        "",
        "2026-11-06",
        "2026-11-06",
      );
    } finally {
      db.close();
    }
    await dialog.getByRole("button", { name: "重新檢查", exact: true }).click();
    await expect(dialog).toContainText("已啟用 1 家公司");
    for (const [button, tab] of [
      ["設定檢查排程", "監控排程"],
      ["設定背景啟動", "啟動與系統匣"],
      ["前往測試通知", "通知"],
    ]) {
      await dialog.getByRole("button", { name: button, exact: true }).click();
      await expect(
        page.getByRole("tab", { name: tab, exact: true }),
      ).toHaveAttribute("aria-selected", "true");
      await page.getByRole("button", { name: "使用指南", exact: true }).click();
      await page
        .getByRole("button", { name: "首次使用設定", exact: true })
        .click();
      await expect(dialog).toContainText("已啟用 1 家公司");
    }
    await page.setViewportSize({ width: 390, height: 844 });
    await expect(dialog).toBeInViewport({ ratio: 1 });
    await page.screenshot({ path: path.join(output, "setup-mobile-top.png") });
    await dialog
      .getByRole("heading", { name: "依工作需要選擇", exact: true })
      .scrollIntoViewIfNeeded();
    await page.screenshot({
      path: path.join(output, "setup-mobile-optional.png"),
    });
    await dialog
      .getByRole("button", { name: "先使用小幫手", exact: true })
      .click();
    await expect(dialog).toHaveCount(0);
    expect(
      await app.evaluate(
        () => (globalThis as unknown as { setupWrites: string[] }).setupWrites,
      ),
    ).toEqual([]);
    expect(
      await page.evaluate(
        () => document.documentElement.scrollWidth <= innerWidth,
      ),
    ).toBe(true);
  } finally {
    await app.evaluate(({ app }) => app.exit(0)).catch(() => undefined);
    await app.close();
    await rm(userData, { recursive: true, force: true });
  }
});
