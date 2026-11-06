import { mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import Database from "better-sqlite3";
import { _electron as electron, expect, test } from "@playwright/test";
import { workbook } from "../support/market-xlsx";

test("market records / daily index state, historical CMoney gap fill, and visible background refresh", async () => {
  test.setTimeout(60_000);
  const root = process.cwd();
  const userData = await mkdtemp(
    path.join(os.tmpdir(), "reporter-daily-index-"),
  );
  const file = path.join(userData, "history-patch.xlsx");
  await writeFile(
    file,
    workbook([
      ["日期", "開盤價", "最高價", "最低價", "收盤價"],
      ["20261103", "120", "130", "108", "125"],
      ["20261104", "888", "999", "125", "888"],
    ]),
  );
  const app = await electron.launch({
    args: [root],
    cwd: root,
    env: {
      ...process.env,
      REPORTER_USER_DATA_DIR: userData,
      REPORTER_TEST_TRAY: "0",
    },
  });
  try {
    const window = await app.firstWindow();
    const db = new Database(path.join(userData, "reporter.sqlite3"));
    try {
      const insert = db.prepare(
        "INSERT INTO market_days(date,kind,data,imported_at) VALUES(?,?,?,?)",
      );
      for (const [date, high, low, close, source] of [
        ["2026-10-30", 110, 90, 100, "CMoney"],
        ["2026-11-02", 120, 90, 110, "CMoney"],
        ["2026-11-04", 155, 125, 150, "TWSE OpenAPI"],
        ["2026-11-05", 165, 145, 160, "TWSE OpenAPI"],
      ] as const)
        insert.run(
          date,
          "index",
          JSON.stringify({ date, high, low, close, source, confirmed: true }),
          "2026-11-06",
        );
      insert.run(
        "2026-11-03",
        "trust",
        JSON.stringify({
          date: "2026-11-03",
          investor: "trust",
          net: 100,
          source: "CMoney",
          confirmed: true,
        }),
        "2026-11-06",
      );
    } finally {
      db.close();
    }
    await app.evaluate(({ dialog }, selected) => {
      dialog.showOpenDialog = (async () => ({
        canceled: false,
        filePaths: [selected],
      })) as typeof dialog.showOpenDialog;
    }, file);
    await window.setViewportSize({ width: 1487, height: 1058 });
    await window.getByRole("tab", { name: "盤面紀錄", exact: true }).click();
    const status = window.getByRole("region", { name: "指數資料更新" });
    await expect(status).toContainText("2026-11-05");
    await expect(status).toContainText("已知指數缺 1 個交易日");
    await expect(status).toContainText("2026-11-03");
    const output = path.join(root, "artifacts/market-records-daily");
    await mkdir(output, { recursive: true });
    await window.screenshot({ path: path.join(output, "missing-history.png") });
    await window.setViewportSize({ width: 390, height: 844 });
    await expect(window.getByRole("button", { name: "匯入 CMoney 補齊" })).toBeInViewport({ ratio: 1 });
    await window.screenshot({ path: path.join(output, "missing-history-mobile.png") });
    await window.getByRole("button", { name: "匯入 CMoney 補齊" }).click();
    await window.getByRole("button", { name: "匯入加權指數 Excel" }).click();
    await expect(window.getByRole("status")).toHaveText("已匯入 2 筆日資料");
    await window.getByRole("button", { name: "完成", exact: true }).click();
    await window.setViewportSize({ width: 1487, height: 1058 });
    await expect(status).not.toContainText("已知指數缺");
    await expect(
      window.getByRole("row").filter({ hasText: "2026-11-04" }),
    ).toContainText("+30.00");
    const saved = new Database(path.join(userData, "reporter.sqlite3"));
    try {
      const row = saved
        .prepare("SELECT data FROM market_days WHERE date=? AND kind='index'")
        .get("2026-11-04") as { data: string };
      expect(JSON.parse(row.data)).toMatchObject({
        high: 155,
        close: 150,
        source: "TWSE OpenAPI",
      });
      saved
        .prepare(
          "INSERT INTO market_days(date,kind,data,imported_at) VALUES(?,?,?,?)",
        )
        .run(
          "2026-11-06",
          "index",
          JSON.stringify({
            date: "2026-11-06",
            high: 190,
            low: 160,
            close: 180,
            source: "TWSE OpenAPI",
            confirmed: true,
          }),
          "2026-11-06",
        );
    } finally {
      saved.close();
    }
    // The hidden Electron harness models a visible page without taking desktop focus.
    await window.evaluate(() => {
      Object.defineProperty(document, "visibilityState", {
        configurable: true,
        value: "visible",
      });
      window.dispatchEvent(new Event("focus"));
    });
    await expect(status).toContainText("2026-11-06");
    await window.screenshot({ path: path.join(output, "history-filled.png") });
  } finally {
    await app.evaluate(({ app }) => app.exit(0)).catch(() => undefined);
    await app.close();
    await rm(userData, { recursive: true, force: true });
  }
});
