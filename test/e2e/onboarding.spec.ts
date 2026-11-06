import { mkdtemp, mkdir, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { _electron as electron, expect, test } from "@playwright/test";

async function launch(userData: string) {
  return electron.launch({
    args: [process.cwd()],
    cwd: process.cwd(),
    env: {
      ...process.env,
      REPORTER_USER_DATA_DIR: userData,
      REPORTER_TEST_TRAY: "0",
    },
  });
}
async function exit(app: Awaited<ReturnType<typeof launch>>) {
  await app.evaluate(({ app }) => app.exit(0)).catch(() => undefined);
  await app.close();
}
async function protectBusinessActions(app: Awaited<ReturnType<typeof launch>>) {
  await app.evaluate(({ ipcMain }) => {
    const calls: string[] = [];
    (globalThis as unknown as { guideWrites: string[] }).guideWrites = calls;
    const channels = [
      "market-records:refresh",
      "market-records:import",
      "calendar:sync",
      "calendar:connect",
      "calendar:disconnect",
      "calendar:import-credentials",
      "calendar:migrate-legacy-token",
      "material-events:soft-delete",
      "material-events:mark-read",
      "schedule:save-settings",
      "schedule:run-now",
      "notification:send-test",
      "notification:schedule-test",
      "data:export",
      "desktop:save-login-startup",
      "watchlist:add",
      "watchlist:update",
      "watchlist:set-active",
      "watchlist:remove",
      "watchlist:import-apply",
      "conference-task:ensure",
      "conference-task:switch-disabled-legacy",
      "conference-task:confirm-legacy",
    ];
    for (const channel of channels) {
      ipcMain.removeHandler(channel);
      ipcMain.handle(channel, () => {
        calls.push(channel);
        throw Error("unexpected business action");
      });
    }
  });
}

test("onboarding / first welcome, skip, real renderer help and persisted completion", async () => {
  test.setTimeout(120_000);
  const userData = await mkdtemp(path.join(os.tmpdir(), "reporter-guide-"));
  let app = await launch(userData);
  try {
    let page = await app.firstWindow();
    await page.setViewportSize({ width: 1487, height: 1058 });
    await expect(
      page.getByRole("region", { name: "歡迎使用小幫手" }),
    ).toBeVisible();
    const output = path.join(process.cwd(), "artifacts/onboarding");
    await mkdir(output, { recursive: true });
    await page.screenshot({ path: path.join(output, "welcome-desktop.png") });
    await page.getByRole("button", { name: "稍後", exact: true }).click();
    await exit(app);
    app = await launch(userData);
    page = await app.firstWindow();
    await expect(
      page.getByRole("heading", { level: 1, name: "總覽", exact: true }),
    ).toBeVisible();
    await expect(
      page.getByRole("region", { name: "歡迎使用小幫手" }),
    ).toHaveCount(0);
    await page.getByRole("button", { name: "使用指南", exact: true }).click();
    await page.getByRole("button", { name: "重看快速導覽" }).click();
    const tour = page.getByRole("dialog", { name: "功能導覽" });
    for (const [i, title] of [
      "關注公司",
      "重大訊息",
      "違約交割",
      "法說會",
      "設定",
    ].entries()) {
      await expect(
        page.getByRole("heading", { level: 1, name: title, exact: true }),
      ).toBeVisible();
      await expect(tour).toContainText(`第 ${i + 1} / 5 步`);
      const targets = [
        '[data-guide="page-actions"]',
        ".list-toolbar",
        ".disclosure-result",
        ".calendar-connection",
        '[role="tablist"][aria-label="設定分類"]',
      ];
      const target = page.locator(targets[i]);
      await expect(target).toHaveClass(/driver-active-element/);
      await expect(target).toBeInViewport({ ratio: 1 });
      await tour
        .getByRole("button", {
          name: i === 4 ? "完成導覽" : "下一步",
          exact: true,
        })
        .click();
    }
    await expect(tour).toHaveCount(0);
    await expect(
      page.getByRole("button", { name: "使用說明", exact: true }),
    ).toBeFocused();
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            JSON.parse(localStorage.getItem("reporter.guided-onboarding.v1")!)
              .status,
        ),
      )
      .toBe("complete");
    await exit(app);
    app = await launch(userData);
    page = await app.firstWindow();
    await expect(
      page.getByRole("heading", { level: 1, name: "總覽", exact: true }),
    ).toBeVisible();
    await expect(
      page.getByRole("region", { name: "歡迎使用小幫手" }),
    ).toHaveCount(0);
  } finally {
    await exit(app);
    await rm(userData, { recursive: true, force: true });
  }
});

test("onboarding / market help, seven page tours, keyboard, scroll and narrow layout without business writes", async () => {
  test.setTimeout(120_000);
  const userData = await mkdtemp(
    path.join(os.tmpdir(), "reporter-guide-layout-"),
  );
  const app = await launch(userData);
  try {
    const page = await app.firstWindow();
    await page.setViewportSize({ width: 1487, height: 1058 });
    await page.emulateMedia({ reducedMotion: "reduce" });
    await page.getByRole("button", { name: "稍後", exact: true }).click();
    await protectBusinessActions(app);
    const output = path.join(process.cwd(), "artifacts/onboarding");
    await mkdir(output, { recursive: true });
    await page.getByRole("tab", { name: "盤面紀錄", exact: true }).click();
    await page.getByRole("button", { name: "使用說明", exact: true }).click();
    const help = page.getByRole("dialog", { name: "盤面紀錄使用說明" });
    await expect(help).toContainText("CMoney");
    await expect(help).toContainText("2020");
    await page.screenshot({
      path: path.join(output, "market-help-desktop.png"),
    });
    await help.getByRole("button", { name: "帶我操作" }).click();
    const tour = page.getByRole("dialog", { name: "功能導覽" });
    await expect(tour).toContainText("第 1 / 3 步");
    await expect(page.locator('[data-guide="page-header"]')).toHaveClass(
      /driver-active-element/,
    );
    await tour.getByRole("button", { name: "下一步", exact: true }).click();
    await expect(tour).toContainText("第 2 / 3 步");
    await tour.getByRole("button", { name: "上一步", exact: true }).click();
    await expect(tour).toContainText("第 1 / 3 步");
    for (let i = 0; i < 6; i++) {
      await page.keyboard.press("Tab");
      expect(
        await tour.evaluate((el) => el.contains(document.activeElement)),
      ).toBe(true);
    }
    await page.screenshot({
      path: path.join(output, "market-tour-desktop.png"),
    });
    await page.keyboard.press("Escape");
    await expect(tour).toHaveCount(0);
    await expect(
      page.getByRole("button", { name: "使用說明", exact: true }),
    ).toBeFocused();
    await page.getByRole("tab", { name: "盤面紀錄", exact: true }).click();
    await page.setViewportSize({ width: 390, height: 844 });
    await page.locator(".reporter-main").evaluate((el) => {
      el.scrollTop = el.scrollHeight;
    });
    await page.getByRole("button", { name: "開啟導覽", exact: true }).click();
    await page.getByRole("button", { name: "使用指南", exact: true }).click();
    await page
      .getByRole("dialog", { name: "使用指南", exact: true })
      .getByRole("button", { name: "盤面紀錄", exact: true })
      .click();
    await expect(help).toBeInViewport();
    expect(
      (await help
        .getByRole("button", { name: "帶我操作", exact: true })
        .boundingBox())!.height,
    ).toBeLessThan(64);
    await page.screenshot({
      path: path.join(output, "market-help-mobile.png"),
    });
    await help.getByRole("button", { name: "帶我操作" }).click();
    for (let i = 0; i < 3; i++) {
      await expect(tour).toContainText(`第 ${i + 1} / 3 步`);
      await expect(tour).toBeInViewport({ ratio: 1 });
      await expect(
        tour.getByRole("button", {
          name: i === 2 ? "完成導覽" : "下一步",
          exact: true,
        }),
      ).toBeInViewport({ ratio: 1 });
      expect(
        await page.evaluate(
          () => document.documentElement.scrollWidth <= innerWidth,
        ),
      ).toBe(true);
      if (i === 0) {
        await expect(page.locator('[data-guide="page-header"]')).toHaveClass(
          /driver-active-element/,
        );
        await expect(page.locator('[data-guide="page-header"]')).toBeInViewport(
          { ratio: 1 },
        );
        expect(
          await page
            .locator("body")
            .evaluate((el) => el.classList.contains("driver-fade")),
        ).toBe(false);
        await page.screenshot({
          path: path.join(output, "market-tour-mobile.png"),
        });
      }
      await tour
        .getByRole("button", {
          name: i === 2 ? "完成導覽" : "下一步",
          exact: true,
        })
        .click();
    }
    expect(
      await app.evaluate(
        () => (globalThis as unknown as { guideWrites: string[] }).guideWrites,
      ),
    ).toEqual([]);
    await expect(page.locator(".driver-overlay")).toHaveCount(0);
  } finally {
    await exit(app);
    await rm(userData, { recursive: true, force: true });
  }
});

test("onboarding / seven page tours use explanations without business writes", async () => {
  test.setTimeout(120_000);
  const userData = await mkdtemp(
    path.join(os.tmpdir(), "reporter-guide-pages-"),
  );
  const app = await launch(userData);
  try {
    const page = await app.firstWindow();
    await page.setViewportSize({ width: 1487, height: 1058 });
    await page.emulateMedia({ reducedMotion: "reduce" });
    await page.getByRole("button", { name: "稍後", exact: true }).click();
    await protectBusinessActions(app);
    const tour = page.getByRole("dialog", { name: "功能導覽" });
    for (const title of [
      "總覽",
      "關注公司",
      "重大訊息",
      "違約交割",
      "法說會",
      "盤面紀錄",
      "設定",
    ]) {
      await page.getByRole("tab", { name: title, exact: true }).click();
      await page.getByRole("button", { name: "使用說明", exact: true }).click();
      await page.getByRole("button", { name: "帶我操作" }).click();
      for (let i = 0; i < 3; i++) {
        await expect(tour).toContainText(`第 ${i + 1} / 3 步`);
        await tour
          .getByRole("button", {
            name: i === 2 ? "完成導覽" : "下一步",
            exact: true,
          })
          .click();
      }
    }

    expect(
      await app.evaluate(
        () => (globalThis as unknown as { guideWrites: string[] }).guideWrites,
      ),
    ).toEqual([]);
    await expect(page.locator(".driver-overlay")).toHaveCount(0);
  } finally {
    await exit(app);
    await rm(userData, { recursive: true, force: true });
  }
});
