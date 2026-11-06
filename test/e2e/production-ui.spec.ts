import { mkdtemp, mkdir, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import Database from "better-sqlite3";
import { _electron as electron, expect, test } from "@playwright/test";

test("production UI / real persisted data / approved seven-page layout and watchlist writes", async () => {
  test.setTimeout(90_000);
  const root = process.cwd();
  const userData = await mkdtemp(
    path.join(os.tmpdir(), "reporter-production-ui-"),
  );
  const launch = () =>
    electron.launch({
      args: [root],
      cwd: root,
      env: {
        ...process.env,
        REPORTER_USER_DATA_DIR: userData,
        REPORTER_TEST_TRAY: "0",
      },
    });
  let app = await launch();
  try {
    let window = await app.firstWindow();
    const db = new Database(path.join(userData, "reporter.sqlite3"));
    try {
      const companies = [
        ["2330", "台積電", "半導體"],
        ["2317", "鴻海", "電子"],
        ["2454", "聯發科", "半導體"],
        ["2881", "富邦金", "金融"],
      ];
      companies.forEach(([code, name, category], i) => {
        db.prepare(
          "INSERT INTO companies(id,market,stock_code,name,updated_at) VALUES(?,?,?,?,?)",
        ).run(code, "TWSE", code, name, new Date().toISOString());
        db.prepare(
          "INSERT INTO watchlist_entries(id,company_id,active,category,notes,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",
        ).run(
          code,
          code,
          1,
          category,
          "",
          new Date().toISOString(),
          new Date().toISOString(),
        );
        db.prepare(
          "INSERT INTO material_events(id,company_id,source_key,content_fingerprint,title,content,published_at) VALUES(?,?,?,?,?,?,?)",
        ).run(
          code,
          code,
          code,
          code,
          [
            "董事會通過第三季財務報告",
            "公告本公司董事會決議事項",
            "公告本公司召開法人說明會",
            "公告本公司每月營收",
          ][i],
          "這是隔離測試保存的公告全文，用來驗證正式介面從 SQLite 讀取內容。",
          new Date(Date.now() - i * 3600000).toISOString(),
        );
      });
      const future = new Date(Date.now() + 2 * 86400000).toISOString();
      db.prepare(
        "INSERT INTO conferences(id,company_id,stock_code,company_name,source_key,starts_at,location,content) VALUES(?,?,?,?,?,?,?,?)",
      ).run(
        "conf",
        "2330",
        "2330",
        "台積電",
        "conf",
        future,
        "線上法人說明會",
        "隔離測試場次",
      );
      for (let i = 0; i < 3; i++) {
        const date = `2026-11-0${i + 2}`;
        const index = {
          date,
          high: 30000 + i * 500,
          low: 29000 + i * 100,
          close: 29500 + i * 100,
          confirmed: true,
          source: "TWSE OpenAPI",
        };
        const flow = {
          date,
          investor: "trust",
          net: (i + 1) * 100000000,
          confirmed: true,
          source: "CMoney",
        };
        for (const [kind, data] of [
          ["index", index],
          ["trust", flow],
        ] as const)
          db.prepare(
            "INSERT INTO market_days(date,kind,data,imported_at) VALUES(?,?,?,?)",
          ).run(date, kind, JSON.stringify(data), new Date().toISOString());
      }
    } finally {
      db.close();
    }
    const output = path.join(root, "artifacts/ui-production/screenshots");
    await mkdir(output, { recursive: true });
    await window.setViewportSize({ width: 1487, height: 1058 });
    for (const [page, file] of [
      ["總覽", "overview"],
      ["關注公司", "watchlist"],
      ["重大訊息", "material"],
      ["違約交割", "disclosures"],
      ["法說會", "conferences"],
      ["盤面紀錄", "records-index"],
      ["設定", "settings"],
    ] as const) {
      await window.getByRole("tab", { name: page, exact: true }).click();
      await expect(
        window.getByRole("heading", { name: page, exact: true }),
      ).toBeVisible();
      if (page === "總覽")
        await expect(window.getByText("台積電").first()).toBeVisible();
      if (page === "重大訊息")
        await expect(
          window.getByRole("button", {
            name: "台積電 董事會通過第三季財務報告",
          }),
        ).toBeVisible();
      if (page === "盤面紀錄") {
        await window.getByLabel("盤面資料日期", { exact: true }).fill("2026-11-04");
        await expect(window.getByText("+1,400.00").first()).toBeVisible();
      }
      await window.screenshot({ path: path.join(output, `${file}.png`) });
      if (page === "總覽") {
        await window.setViewportSize({ width: 1672, height: 941 });
        await window.screenshot({ path: path.join(output, "overview-native.png") });
        await window.setViewportSize({ width: 1487, height: 1058 });
      }
      if (page === "重大訊息") {
        await window.getByRole("button", { name: "台積電 董事會通過第三季財務報告" }).click();
        await expect(window.getByRole("complementary", { name: "重大訊息完整內容" })).toContainText("隔離測試保存的公告全文");
        await window.screenshot({ path: path.join(output, "material-expanded.png") });
        await window.getByRole("button", { name: "關閉重大訊息詳情" }).click();
      }
      if (page === "設定") {
        for (const [tab, name] of [["通知", "notifications"], ["啟動與系統匣", "startup"], ["資料管理", "data"]]) {
          await window.getByRole("tab", { name: tab, exact: true }).click();
          await window.screenshot({ path: path.join(output, `settings-${name}.png`) });
        }
      }
    }
    await window.getByRole("tab", { name: "盤面紀錄", exact: true }).click();
    await window.getByRole("tab", { name: "三大法人", exact: true }).click();
    await window.getByRole("tab", { name: "投信", exact: true }).click();
    await expect(
      window.getByRole("row").filter({ hasText: "2026-11-02" }),
    ).toContainText("6.00");
    await window.screenshot({
      path: path.join(output, "records-investors.png"),
    });
    await window.getByRole("button", { name: "資料管理", exact: true }).click();
    await expect(
      window.getByRole("dialog", { name: "盤面資料管理" }),
    ).toBeVisible();
    await window.screenshot({ path: path.join(output, "import.png") });
    await window.getByRole("button", { name: "完成", exact: true }).click();
    await window.setViewportSize({ width: 390, height: 844 });
    await window.getByRole("button", { name: "開啟導覽" }).click();
    await window.getByRole("tab", { name: "重大訊息", exact: true }).click();
    await expect(
      window.getByRole("button", { name: "台積電 董事會通過第三季財務報告" }),
    ).toBeVisible();
    expect(
      await window.evaluate(
        () => document.documentElement.scrollWidth - window.innerWidth,
      ),
    ).toBeLessThanOrEqual(1);
    await window.screenshot({ path: path.join(output, "material-mobile.png") });
    await window.setViewportSize({ width: 1487, height: 1058 });
    await window.getByRole("tab", { name: "關注公司", exact: true }).click();
    await window.getByRole("button", { name: "編輯台積電" }).click();
    await window
      .getByRole("textbox", { name: "分類", exact: true })
      .fill("測試分類");
    await window
      .getByRole("textbox", { name: "備註", exact: true })
      .fill("真實 IPC 持久化測試");
    await window.getByRole("button", { name: "儲存", exact: true }).click();
    await expect(window.getByRole("dialog")).toHaveCount(0);
    await expect(
      window.getByRole("row").filter({ hasText: "台積電" }),
    ).toContainText("測試分類");
    await window.getByRole("switch", { name: "監控台積電" }).click();
    await expect(
      window.getByRole("switch", { name: "監控台積電" }),
    ).toHaveAttribute("aria-checked", "false");
    await app.evaluate(({ app }) => app.exit(0));
    await app.close();
    app = await launch();
    window = await app.firstWindow();
    await window.getByRole("tab", { name: "關注公司", exact: true }).click();
    await expect(
      window.getByRole("switch", { name: "監控台積電" }),
    ).toHaveAttribute("aria-checked", "false");
    await expect(
      window.getByRole("row").filter({ hasText: "台積電" }),
    ).toContainText("測試分類");
    const saved = new Database(path.join(userData, "reporter.sqlite3"), {
      readonly: true,
    });
    try {
      expect(
        saved
          .prepare(
            "SELECT active,category,notes FROM watchlist_entries WHERE company_id=?",
          )
          .get("2330"),
      ).toEqual({
        active: 0,
        category: "測試分類",
        notes: "真實 IPC 持久化測試",
      });
    } finally {
      saved.close();
    }
  } finally {
    await app.evaluate(({ app }) => app.exit(0)).catch(() => undefined);
    await app.close();
    await rm(userData, { recursive: true, force: true });
  }
});
