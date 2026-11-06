import { mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { _electron as electron, expect, test } from "@playwright/test";
test("market-records / isolated Electron imports CSV and preserves a source error", async () => {
  test.setTimeout(60_000);
  const root = process.cwd(),
    userData = await mkdtemp(path.join(os.tmpdir(), "market-records-e2e-"));
  const csv = path.join(userData, "daily.csv");
  const earlier = path.join(userData, "earlier.csv");
  await writeFile(
    earlier,
    "115年11月03日 三大法人買賣金額統計表\n自營商(自行買賣),100000000,0,100000000\n自營商(避險),50000000,0,50000000\n投信,300000000,0,300000000\n外資及陸資(不含外資自營商),100,0,100",
  );
  await writeFile(
    csv,
    "115年11月04日 三大法人買賣金額統計表\n自營商(自行買賣),100,0,100\n自營商(避險),100,0,100\n投信,200000000,0,200000000\n外資及陸資(不含外資自營商),100,0,100",
  );
  let app = await electron.launch({
    args: [root],
    cwd: root,
    env: {
      ...process.env,
      REPORTER_USER_DATA_DIR: userData,
      REPORTER_TEST_TRAY: "0",
    },
  });
  try {
    let window = await app.firstWindow();
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = (async () => ({
        canceled: false,
        filePaths: [file],
      })) as typeof dialog.showOpenDialog;
    }, csv);
    await window.getByRole("tab", { name: "盤面紀錄" }).click();
    await expect(
      window.getByRole("heading", { name: "盤面紀錄", exact: true }),
    ).toBeVisible();
    await window.getByRole("button", { name: "資料管理", exact: true }).click();
    await window.getByRole("button", { name: "匯入 TWSE 當日 CSV" }).click();
    await expect(window.getByRole("status")).toHaveText("已匯入 3 筆日資料");
    await window.getByRole("button", { name: "完成", exact: true }).click();
    await expect(
      window.getByText(/資料範圍 2026-11-04—2026-11-04/),
    ).toBeVisible();
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = (async () => ({
        canceled: false,
        filePaths: [file],
      })) as typeof dialog.showOpenDialog;
    }, earlier);
    await window.getByRole("button", { name: "資料管理", exact: true }).click();
    await window.getByRole("button", { name: "匯入 TWSE 當日 CSV" }).click();
    await expect(window.getByRole("status")).toHaveText("已匯入 3 筆日資料");
    await window.getByRole("button", { name: "完成", exact: true }).click();
    await window.getByRole("tab", { name: "三大法人", exact: true }).click();
    await window.getByRole("tab", { name: "投信", exact: true }).click();
    const buyRow = window.getByRole("row").filter({ hasText: "2026-11-03" });
    await expect(buyRow).toContainText("5.00");
    await expect(buyRow).toContainText("01");
    await expect(buyRow).toContainText("2 日");
    await app.evaluate(({ app }) => app.exit(0));
    await app.close();
    app = await electron.launch({
      args: [root],
      cwd: root,
      env: {
        ...process.env,
        REPORTER_USER_DATA_DIR: userData,
        REPORTER_TEST_TRAY: "0",
      },
    });
    window = await app.firstWindow();
    await window.getByRole("tab", { name: "盤面紀錄", exact: true }).click();
    await window.getByRole("tab", { name: "三大法人", exact: true }).click();
    await window.getByRole("tab", { name: "投信", exact: true }).click();
    await expect(
      window.getByRole("row").filter({ hasText: "2026-11-03" }),
    ).toContainText("5.00");
    await window.getByRole("button", { name: "更新官方指數" }).click();
    await expect(window.getByRole("alert")).toContainText(
      "隔離預覽不連正式資料來源",
    );
    await window.screenshot({
      path: path.join(root, "artifacts", "market-records-preview.png"),
      fullPage: true,
    });
  } finally {
    await app.close();
    await rm(userData, { recursive: true, force: true });
  }
});
