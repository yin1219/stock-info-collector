import { mkdtemp, mkdir, rm } from "node:fs/promises";
import { createServer } from "node:http";
import os from "node:os";
import path from "node:path";
import { _electron as electron, expect, test } from "@playwright/test";

test("header layout / local metadata stays clear of primary actions at top and while scrolling", async () => {
  test.setTimeout(90_000);
  const root = process.cwd();
  const userData = await mkdtemp(
    path.join(os.tmpdir(), "reporter-header-layout-"),
  );
  const requests: string[] = [];
  const server = createServer((request, response) => {
    requests.push(request.url ?? "");
    response.statusCode = 503;
    response.end("Offline header test fixture");
  });
  await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
  const address = server.address();
  if (!address || typeof address === "string")
    throw new Error("Cannot start header fixture");
  let app: Awaited<ReturnType<typeof electron.launch>> | undefined;
  try {
    app = await electron.launch({
      args: [root],
      cwd: root,
      env: {
        ...process.env,
        REPORTER_USER_DATA_DIR: userData,
        REPORTER_TEST_TRAY: "0",
        REPORTER_TEST_LIVE_SOURCES: "1",
        REPORTER_TEST_HTTP_PROXY: `http://127.0.0.1:${address.port}/source`,
      },
    });
    const window = await app.firstWindow();
    const output = path.join(root, "artifacts/ui-header-layout");
    await mkdir(output, { recursive: true });
    for (const { width, pages } of [
      { width: 1200, pages: ["重大訊息", "關注公司", "盤面紀錄", "設定"] },
      { width: 960, pages: ["重大訊息"] },
      { width: 1487, pages: ["重大訊息"] },
      { width: 390, pages: ["重大訊息", "設定"] },
    ]) {
      await window.setViewportSize({
        width,
        height: width === 390 ? 844 : 900,
      });
      for (const page of pages) {
        if (width <= 700)
          await window.getByRole("button", { name: "開啟導覽" }).click();
        await window.getByRole("tab", { name: page, exact: true }).click();
        await expect(
          window.getByRole("heading", { name: page, exact: true }),
        ).toBeVisible();
        await window.locator(".reporter-main").evaluate((main) => {
          main.scrollTop = 0;
        });
        const layout = await window
          .locator(".header-actions")
          .evaluate((actions) => {
            const label = document
              .querySelector(".demo-label")!
              .getBoundingClientRect();
            const heading = document
              .querySelector("h1")!
              .getBoundingClientRect();
            const intersects = (rect: DOMRect) =>
              Math.min(rect.right, label.right) >
                Math.max(rect.left, label.left) &&
              Math.min(rect.bottom, label.bottom) >
                Math.max(rect.top, label.top);
            return {
              headingOverlaps: intersects(heading),
              controlsOverlap: [
                ...actions.querySelectorAll("button,input,select"),
              ].some((control) => intersects(control.getBoundingClientRect())),
            };
          });
        if (page === "重大訊息")
          await window.screenshot({
            path: path.join(output, `material-${width}.png`),
          });
        expect(layout, `${page} at ${width}px`).toEqual({
          headingOverlaps: false,
          controlsOverlap: false,
        });
        if (page === "重大訊息") {
          const check = window.getByRole("button", {
            name: "立即檢查重大訊息",
            exact: true,
          });
          await expect(check).toBeEnabled();
          await check.click({ trial: true });
        }
        if (page === "設定") {
          await window
            .getByRole("button", { name: "儲存變更", exact: true })
            .click({ trial: true });
          await window.locator(".reporter-main").evaluate((main) => {
            main.scrollTop = 32;
          });
          const hit = await window
            .getByRole("button", { name: "儲存變更", exact: true })
            .evaluate((button) => {
              const rect = button.getBoundingClientRect();
              return [
                [8, 8],
                [rect.width / 2, rect.height / 2],
                [rect.width - 8, 8],
              ].map(([x, y]) => {
                const element = document.elementFromPoint(
                  rect.left + x,
                  rect.top + y,
                );
                return {
                  inside:
                    element === button ||
                    (element !== null && button.contains(element)),
                  x: rect.left + x,
                  y: rect.top + y,
                  target: element?.outerHTML.slice(0, 160),
                };
              });
            });
          await window.screenshot({
            path: path.join(output, `settings-scrolled-${width}.png`),
          });
          expect(
            hit.every((point) => point.inside),
            `save action at ${width}px after scrolling ${JSON.stringify(hit)}`,
          ).toBe(true);
          await window.locator(".reporter-main").evaluate((main) => {
            main.scrollTop = 0;
          });
        }
      }
    }
    expect(requests).toEqual([]);
  } finally {
    if (app) {
      await app.evaluate(({ app }) => app.exit(0)).catch(() => undefined);
      await app.close();
    }
    await new Promise<void>((resolve) => server.close(() => resolve()));
    await rm(userData, { recursive: true, force: true });
  }
});
