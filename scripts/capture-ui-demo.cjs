// Capture only the isolated browser demo; never launch Electron or read userData.
const { chromium } = require("playwright");
const fs = require("node:fs/promises");
const path = require("node:path");
(async () => {
  const output = path.resolve("artifacts/ui-demo/screenshots");
  await fs.mkdir(output, { recursive: true });
  const browser = await chromium.launch({
    channel: process.env.REPORTER_DEMO_BROWSER || "chrome",
  });
  try {
    const page = await browser.newPage({
      viewport: { width: 1487, height: 1058 },
      reducedMotion: "reduce",
    });
    await page.route("**/*", (route) =>
      /^(127\.0\.0\.1|localhost)$/.test(new URL(route.request().url()).hostname)
        ? route.continue()
        : route.abort(),
    );
    await page.goto("http://127.0.0.1:5178");
    const pages = {
      overview: "總覽",
      watchlist: "關注公司",
      material: "重大訊息",
      disclosures: "違約交割",
      conferences: "法說會",
      "records-index": "盤面紀錄",
      settings: "設定",
    };
    for (const [file, label] of Object.entries(pages)) {
      await page.getByRole("button", { name: label, exact: true }).click();
      if (file === "material")
        await page
          .getByRole("button", { name: "董事會決議股利分配", exact: true })
          .click();
      if (file === "watchlist")
        await page.getByRole("button", { name: "更多鴻海操作" }).click();
      await page.mouse.move(1480, 0);
      await page.screenshot({ path: path.join(output, `${file}.png`) });
    }
    for (const [file, label] of Object.entries({
      "settings-notifications": "通知",
      "settings-startup": "啟動與系統匣",
      "settings-data": "資料管理",
    })) {
      await page.getByRole("tab", { name: label, exact: true }).click();
      await page.screenshot({ path: path.join(output, `${file}.png`) });
    }
    await page.getByRole("button", { name: "盤面紀錄", exact: true }).click();
    await page.getByRole("tab", { name: "三大法人", exact: true }).click();
    await page.screenshot({ path: path.join(output, "records-investors.png") });
    await page.setViewportSize({ width: 390, height: 844 });
    await page.screenshot({ path: path.join(output, "mobile-records.png") });
    await page.getByRole("button", { name: "開啟導覽" }).click();
    await page.getByRole("button", { name: "設定", exact: true }).click();
    await page.screenshot({ path: path.join(output, "mobile-settings.png") });
    await page.reload();
    await page.getByRole('heading',{name:'總覽',level:1,exact:true}).waitFor();
    await page.setViewportSize({ width: 1672, height: 941 });
    await page.getByRole("button", { name: "總覽", exact: true }).click();
    await page.mouse.move(1660, 0);
    await page.screenshot({ path: path.join(output, "overview-native.png") });
    console.log(`Captured 14 demo views: ${output}`);
  } finally {
    await browser.close();
  }
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
