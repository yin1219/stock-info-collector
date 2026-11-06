import { test as base, expect } from "@playwright/test";
const test = base.extend<{ denyExternalRequests: void }>({
  denyExternalRequests: [
    async ({ page }, use) => {
      const external: string[] = [];
      await page.route("**/*", (route) => {
        if (
          !new URL(route.request().url()).hostname.match(
            /^(127\.0\.0\.1|localhost)$/,
          )
        ) {
          external.push(route.request().url());
          return route.abort();
        }
        return route.continue();
      });
      await use();
      expect(external, "demo must not request external services").toEqual([]);
    },
    { auto: true },
  ],
});
test("demo / complete navigation, inline detail and responsive controls", async ({
  page,
}) => {
  await page.goto("/");
  for (const name of [
    "總覽",
    "關注公司",
    "重大訊息",
    "違約交割",
    "法說會",
    "盤面紀錄",
    "設定",
  ]) {
    await page.getByRole("button", { name, exact: true }).click();
    await expect(
      page.getByRole("heading", { level: 1, name, exact: true }),
    ).toBeVisible();
  }
  await page.getByRole("button", { name: "重大訊息", exact: true }).click();
  await page
    .getByRole("button", { name: "董事會決議股利分配", exact: true })
    .click();
  await expect(page.getByText("公告內容", { exact: true })).toBeVisible();
  await page.getByRole("button", { name: "關閉公告內容" }).click();
  await page.getByRole("button", { name: "設定", exact: true }).click();
  await page.getByLabel("檢查頻率").selectOption("30");
  await page.getByRole("button", { name: "儲存變更" }).click();
  await expect(page.getByRole("status")).toContainText("已儲存示意設定");
  await page.setViewportSize({ width: 390, height: 844 });
  await page.getByRole("button", { name: "開啟導覽" }).click();
  await page.getByRole("button", { name: "盤面紀錄", exact: true }).click();
  await expect(
    page.getByRole("heading", { level: 1, name: "盤面紀錄" }),
  ).toBeVisible();
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= innerWidth,
    ),
  ).toBe(true);
});

test("local watchlist edits, source failure and calendar dedup simulation", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByRole("button", { name: "關注公司", exact: true }).click();
  await page.getByRole("button", { name: "新增關注公司" }).click();
  const dialog = page.getByRole("dialog");
  await dialog.getByLabel("股票代號").fill("9999");
  await dialog.getByLabel("公司名稱").fill("測試公司");
  await dialog.getByRole("button", { name: "新增", exact: true }).click();
  await expect(
    page.getByRole("cell", { name: "測試公司", exact: true }),
  ).toBeVisible();
  await page.getByRole("button", { name: "編輯測試公司", exact: true }).click();
  await dialog.getByLabel("公司名稱").fill("編輯完成");
  await dialog.getByRole("button", { name: "儲存", exact: true }).click();
  await page.getByRole("button", { name: "更多編輯完成操作" }).click();
  await page.getByRole("button", { name: "移除編輯完成" }).click();
  await dialog.getByRole("button", { name: "取消", exact: true }).click();
  await expect(
    page.getByRole("cell", { name: "編輯完成", exact: true }),
  ).toBeVisible();
  await page.getByRole("button", { name: "移除編輯完成" }).click();
  await dialog.getByRole("button", { name: "確認移除" }).click();
  await expect(
    page.getByRole("cell", { name: "編輯完成", exact: true }),
  ).toHaveCount(0);
  await page.getByRole("button", { name: "重大訊息", exact: true }).click();
  await page
    .getByRole("button", { name: "董事會決議股利分配", exact: true })
    .click();
  await page.getByRole("button", { name: "刪除此筆本機公告" }).click();
  await dialog.getByRole("button", { name: "確認刪除" }).click();
  await expect(
    page.getByRole("button", { name: "董事會決議股利分配", exact: true }),
  ).toHaveCount(0);
  await page.getByRole("button", { name: "違約交割", exact: true }).click();
  await expect(
    page.getByRole("heading", { name: "當日無個股達違約資訊揭露標準" }),
  ).toBeVisible();
  await page.getByText("切換示意狀態", { exact: true }).click();
  await page.getByLabel("來源情境").selectOption("failed");
  await expect(
    page.getByRole("heading", { name: "今日資料尚未確認" }),
  ).toBeVisible();
  await expect(
    page.getByRole("heading", { name: "當日無個股達違約資訊揭露標準" }),
  ).toHaveCount(0);
  await page.getByRole("button", { name: "法說會", exact: true }).click();
  for (let i = 0; i < 2; i++) {
    await page.getByRole("button", { name: "立即擷取並同步法說會" }).click();
    await expect(page.getByRole("status")).toContainText("3 筆已存在");
  }
  await expect(page.getByRole("table").getByRole("row")).toHaveCount(4);
});

test("whole announcement row and shared CMoney import guide work on desktop and mobile", async ({
  page,
}) => {
  await page.goto("/");
  for (const width of [1487, 390]) {
    await page.setViewportSize({ width, height: 1058 });
    if (width === 390)
      await page.getByRole("button", { name: "開啟導覽" }).click();
    await page.getByRole("button", { name: "重大訊息", exact: true }).click();
    const row = page.getByRole("button", {
      name: "董事會決議股利分配",
      exact: true,
    });
    await row.getByText("11/05 14:32", { exact: true }).click();
    await expect(row).toHaveAttribute("aria-expanded", "true");
    await row.getByText("台積電", { exact: true }).click();
    await expect(row).toHaveAttribute("aria-expanded", "false");
    await row.focus();
    await page.keyboard.press("Enter");
    await expect(page.getByText("公告內容", { exact: true })).toBeVisible();
    await row.click({ position: { x: 4, y: 4 } });
    await expect(row).toHaveAttribute("aria-expanded", "false");
    await row.getByText("台積電", { exact: true }).click();
    await page.getByRole("button", { name: "關閉公告內容" }).click();
    if (width === 390)
      await page.getByRole("button", { name: "開啟導覽" }).click();
    await page.getByRole("button", { name: "盤面紀錄", exact: true }).click();
    for (const tab of ["加權指數", "三大法人"]) {
      await page.getByRole("tab", { name: tab, exact: true }).click();
      await page.getByRole("button", { name: "資料管理", exact: true }).click();
      const dialog = page.getByRole("dialog");
      await expect(dialog.getByText(/不同分頁共用同一份資料/)).toBeVisible();
      await expect(dialog.getByText(/CMoney 預設表格匯出/)).toBeVisible();
      await expect(dialog.getByText(/不必每次重匯完整歷史/)).toBeVisible();
      await expect(dialog.getByText(/2020年/)).toHaveCount(0);
      await expect(dialog.locator('input[accept=".xlsx"]')).toHaveCount(4);
      expect(
        await dialog.evaluate((el) => el.scrollWidth <= el.clientWidth),
      ).toBe(true);
      await dialog
        .getByRole("heading", { name: "自營商", level: 4, exact: true })
        .scrollIntoViewIfNeeded();
      await expect(
        dialog.getByRole("heading", { name: "自營商", level: 4, exact: true }),
      ).toBeVisible();
      if (tab === "加權指數") {
        await dialog.evaluate((el) => {
          el.scrollTop = 0;
        });
        await page.screenshot({
          path: `artifacts/ui-demo/screenshots/import-guide-${width}.png`,
        });
      }
      await dialog.getByRole("button", { name: "完成", exact: true }).click();
    }
  }
});

test("settings persist between tabs, export is demo-only, and mobile pages fit", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByRole("button", { name: "設定", exact: true }).click();
  await page.getByLabel("檢查頻率").selectOption("30");
  await page.getByRole("tab", { name: "通知", exact: true }).click();
  await page.getByRole("button", { name: "一分鐘後發送測試通知" }).click();
  await page.getByRole("button", { name: "取消排定" }).click();
  await expect(page.getByRole("status")).toContainText("已取消");
  await page.getByRole("tab", { name: "啟動與系統匣", exact: true }).click();
  await page.getByRole("switch", { name: "登入 Windows 時啟動" }).click();
  await expect(
    page.getByRole("switch", { name: "登入 Windows 時啟動" }),
  ).toHaveAttribute("aria-checked", "false");
  await page.getByRole("tab", { name: "資料管理", exact: true }).click();
  const downloadEvent = page.waitForEvent("download");
  await page.getByRole("button", { name: "匯出示意設定" }).click();
  const download = await downloadEvent;
  expect(download.suggestedFilename()).toBe("reporter-demo-settings.json");
  await page.getByRole("tab", { name: "監控排程", exact: true }).click();
  await expect(page.getByLabel("檢查頻率")).toHaveValue("30");
  await page.getByRole("tab", { name: "資料管理", exact: true }).click();
  await page.getByRole("button", { name: "重置示意資料" }).click();
  await page
    .getByRole("dialog")
    .getByRole("button", { name: "確認重置" })
    .click();
  await expect(page.getByLabel("檢查頻率")).toHaveValue("15");
  await page.getByRole("tab", { name: "啟動與系統匣", exact: true }).click();
  await expect(
    page.getByRole("switch", { name: "登入 Windows 時啟動" }),
  ).toHaveAttribute("aria-checked", "true");
  await page.setViewportSize({ width: 390, height: 844 });
  for (const name of [
    "總覽",
    "關注公司",
    "重大訊息",
    "違約交割",
    "法說會",
    "盤面紀錄",
    "設定",
  ]) {
    await page.getByRole("button", { name: "開啟導覽" }).click();
    await page.getByRole("button", { name, exact: true }).click();
    await expect(
      page.getByRole("heading", { name, level: 1, exact: true }),
    ).toBeVisible();
    expect(
      await page.evaluate(
        () => document.documentElement.scrollWidth <= innerWidth,
      ),
      name,
    ).toBe(true);
  }
  await page.getByRole("button", { name: "開啟導覽" }).click();
  await page.getByRole("button", { name: "關注公司", exact: true }).click();
  const table = page.locator(".table-scroll");
  expect(await table.evaluate((el) => el.scrollWidth > el.clientWidth)).toBe(
    true,
  );
  await table.evaluate((el) => {
    el.scrollLeft = el.scrollWidth;
  });
  expect(await table.evaluate((el) => el.scrollLeft)).toBeGreaterThan(0);
  await page.getByRole("button", { name: "編輯台積電", exact: true }).click();
  await expect(page.getByRole("dialog").getByLabel("公司名稱")).toHaveValue(
    "台積電",
  );
  await page.keyboard.press("Escape");
  await expect(page.getByRole("dialog")).toHaveCount(0);
});
