import { describe, it, expect, vi } from "vitest";
import { openDatabase } from "../../src/repositories/migrations";
import { createMarketRecordsStore } from "../../src/repositories/market-records";
import { createMarketRecordsService } from "../../src/services/market-records";
import { startSchedulerLifecycle } from "../../src/main/scheduler-bootstrap";
import { createTestPaths } from "../support";
import { workbook } from "../support/market-xlsx";

describe("market-records daily lifecycle", () => {
  it("updates after close and resume, retains history through failures and month rollover, and reopens SQLite", async () => {
    const paths = await createTestPaths();
    let now = new Date("2026-10-30T07:29:00Z");
    const firstMonth = [
      {
        Date: "1151030",
        HighestIndex: "120",
        LowestIndex: "90",
        ClosingIndex: "110",
      },
    ];
    const fetchIndex = vi.fn().mockResolvedValue(firstMonth);
    const notify = vi.fn(() => true);
    const db = openDatabase(paths.database);
    try {
      const store = createMarketRecordsStore(db);
      const service = createMarketRecordsService({
        db,
        now: () => now,
        fetchIndex,
        notify,
      });
      service.importFile(
        workbook([
          ["日期", "開盤價", "最高價", "最低價", "收盤價"],
          ["20261029", "100", "105", "95", "100"],
        ]),
        "index",
      );
      const callbacks: Array<() => void> = [];
      let resume: (() => void) | undefined;
      const removeListener = vi.fn();
      const clear = vi.fn();
      const timer = 17 as unknown as ReturnType<typeof setInterval>;
      const lifecycle = startSchedulerLifecycle({
        isolated: false,
        createScheduler: () => ({ tick: service.tick, catchUp: service.tick }),
        powerMonitor: {
          on: (_event, listener) => {
            resume = listener;
          },
          removeListener,
        },
        setInterval: (callback, milliseconds) => {
          expect(milliseconds).toBe(60_000);
          callbacks.push(callback);
          return timer;
        },
        clearInterval: clear,
      })!;
      expect(fetchIndex).not.toHaveBeenCalled();
      now = new Date("2026-10-30T07:30:00Z");
      callbacks[0]();
      await vi.waitFor(() =>
        expect(service.snapshot().indexUpdate.latestOfficialDate).toBe(
          "2026-10-30",
        ),
      );
      expect(fetchIndex).toHaveBeenCalledOnce();
      now = new Date("2026-10-30T07:31:00Z");
      callbacks[0]();
      expect(fetchIndex).toHaveBeenCalledOnce();
      const saved = store.indices();
      fetchIndex.mockRejectedValueOnce(new Error("fixture offline"));
      now = new Date("2026-10-30T08:30:00Z");
      callbacks[0]();
      await vi.waitFor(() =>
        expect(service.snapshot().error).toContain("fixture offline"),
      );
      expect(store.indices()).toEqual(saved);
      fetchIndex.mockResolvedValue([
        {
          Date: "1151102",
          HighestIndex: "150",
          LowestIndex: "100",
          ClosingIndex: "130",
        },
      ]);
      now = new Date("2026-11-02T07:30:00Z");
      resume!();
      await vi.waitFor(() =>
        expect(service.snapshot().indexUpdate.latestOfficialDate).toBe(
          "2026-11-02",
        ),
      );
      expect(store.indices().map((row) => row.date)).toEqual([
        "2026-10-29",
        "2026-10-30",
        "2026-11-02",
      ]);
      expect(service.snapshot().indices.rise).toContainEqual(
        expect.objectContaining({ date: "2026-11-02", amount: 40 }),
      );
      lifecycle.stop();
      expect(clear).toHaveBeenCalledWith(timer);
      expect(removeListener).toHaveBeenCalledWith("resume", resume);
      db.close();
      const reopened = openDatabase(paths.database);
      try {
        const secondNotify = vi.fn(() => true);
        const second = createMarketRecordsService({
          db: reopened,
          now: () => now,
          fetchIndex,
          notify: secondNotify,
        });
        expect(second.snapshot().indexUpdate.latestOfficialDate).toBe(
          "2026-11-02",
        );
        await second.tick();
        expect(secondNotify).not.toHaveBeenCalled();
        expect(createMarketRecordsStore(reopened).indices()).toHaveLength(3);
      } finally {
        reopened.close();
      }
    } finally {
      if (db.open) db.close();
      await paths.cleanup();
    }
  });
});
