import { describe, it, expect, vi } from "vitest";
import { openDatabase } from "../../src/repositories/migrations";
import { createMarketRecordsService } from "../../src/services/market-records";
import { workbook } from "../support/market-xlsx";
import { createMarketRecordsStore } from "../../src/repositories/market-records";
describe("market-records service", () => {
  it("reports official index updates and lets CMoney fill a known missing trading close", async () => {
    const db = openDatabase(":memory:");
    const now = new Date("2026-11-06T08:00:00Z");
    const fetchIndex = vi.fn().mockResolvedValue([
      {
        Date: "1151104",
        HighestIndex: "155",
        LowestIndex: "125",
        ClosingIndex: "150",
      },
      {
        Date: "1151105",
        HighestIndex: "165",
        LowestIndex: "145",
        ClosingIndex: "160",
      },
    ]);
    try {
      const service = createMarketRecordsService({
        db,
        now: () => now,
        fetchIndex,
        notify: () => false,
      });
      service.importFile(
        workbook([
          ["日期", "開盤價", "最高價", "最低價", "收盤價"],
          ["20261030", "100", "110", "90", "100"],
          ["20261102", "110", "120", "90", "110"],
          ["20261104", "140", "150", "100", "140"],
        ]),
        "index",
      );
      service.importFile(
        workbook([
          ["日期", "上市買賣超金額(百萬)"],
          ["20261102", "1"],
          ["20261103", "1"],
          ["20261104", "1"],
        ]),
        "trust",
      );
      expect(service.snapshot().indexUpdate).toMatchObject({
        automatic: true,
        latestOfficialDate: null,
        checkedAt: null,
        knownMissingDates: ["2026-11-03"],
      });
      expect(
        service
          .snapshot()
          .indices.rise.some((row) => row.date === "2026-11-04"),
      ).toBe(false);
      await service.refresh();
      expect(service.snapshot().indexUpdate).toMatchObject({
        latestOfficialDate: "2026-11-05",
        checkedAt: now.toISOString(),
        knownMissingDates: ["2026-11-03"],
      });
      service.importFile(
        workbook([
          ["日期", "開盤價", "最高價", "最低價", "收盤價"],
          ["20261103", "120", "130", "108", "125"],
          ["20261104", "888", "999", "125", "888"],
        ]),
        "index",
      );
      expect(service.snapshot().indexUpdate.knownMissingDates).toEqual([]);
      expect(service.snapshot().indices.rise).toContainEqual(
        expect.objectContaining({ date: "2026-11-04", amount: 30 }),
      );
      const stored = createMarketRecordsStore(db).indices();
      expect(stored).toHaveLength(5);
      expect(stored).toContainEqual(
        expect.objectContaining({ date: "2026-10-30", source: "CMoney" }),
      );
      expect(stored).toContainEqual(
        expect.objectContaining({
          date: "2026-11-04",
          high: 155,
          close: 150,
          source: "TWSE OpenAPI",
        }),
      );
      fetchIndex.mockRejectedValueOnce(new Error("fixture unavailable"));
      await service.refresh();
      expect(createMarketRecordsStore(db).indices()).toEqual(stored);
      expect(service.snapshot().error).toContain("fixture unavailable");
      expect(service.snapshot().indexUpdate.latestOfficialDate).toBe(
        "2026-11-05",
      );
    } finally {
      db.close();
    }
  });

  it("disables automatic ticks while preserving explicit fixture refreshes in isolated mode", async () => {
    const db = openDatabase(":memory:");
    const fetchIndex = vi.fn().mockResolvedValue([]);
    try {
      const service = createMarketRecordsService({
        db,
        now: () => new Date("2026-11-06T08:00:00Z"),
        fetchIndex,
        notify: () => false,
        automaticUpdatesEnabled: false,
      });
      await service.tick();
      expect(fetchIndex).not.toHaveBeenCalled();
      expect(service.snapshot().indexUpdate.automatic).toBe(false);
      await service.refresh();
      expect(fetchIndex).toHaveBeenCalledOnce();
    } finally {
      db.close();
    }
  });

  it("provides the latest confirmed streak even when it is below the top three", () => {
    const db = openDatabase(":memory:");
    try {
      const store = createMarketRecordsStore(db);
      const amounts = [15, 15, -1, 10, 10, -1, 5, 5, -1, 1, 1];
      store.importRows(
        amounts.map((net, i) => ({
          date: `2026-10-${String(i + 1).padStart(2, "0")}`,
          investor: "trust" as const,
          net,
          confirmed: true,
        })),
        [],
        "CMoney",
        "2026-11-05",
      );
      const service = createMarketRecordsService({
        db,
        now: () => new Date("2026-11-05T08:00:00Z"),
        fetchIndex: async () => [],
        notify: () => false,
      });
      const data = service.snapshot();
      expect(
        data.flowGroups.find((g) => g.metric === "trust-buy")?.rows,
      ).toHaveLength(3);
      expect(
        (data as unknown as { latestFlows: Array<unknown> }).latestFlows,
      ).toContainEqual(
        expect.objectContaining({
          investor: "trust",
          dataDate: "2026-10-11",
          direction: "buy",
          streak: expect.objectContaining({ rank: 4, days: 2, amount: 2 }),
        }),
      );
    } finally {
      db.close();
    }
  });
  it("continues daily CSV foreign rankings after three overlapping classification checks", () => {
    const db = openDatabase(":memory:");
    try {
      const s = createMarketRecordsService({
        db,
        now: () => new Date("2026-11-05T08:00:00Z"),
        fetchIndex: async () => [],
        notify: () => false,
      });
      s.importFile(
        workbook([
          ["日期", "上市買賣超金額(百萬)"],
          ["20261102", "1.5"],
          ["20261103", "1.5"],
          ["20261104", "1.5"],
        ]),
        "foreign",
      );
      for (const day of ["02", "03", "04", "05"])
        s.importFile(
          Buffer.from(
            `115年11月${day}日 三大法人買賣金額統計表\n自營商(自行買賣),100,0,100\n自營商(避險),100,0,100\n投信,100,0,100\n外資及陸資(不含外資自營商),1000000,0,1000000\n外資自營商,500000,0,500000`,
          ),
          "csv",
        );
      const data = s.snapshot();
      expect(data.foreignWarning).toBeNull();
      expect(
        data.flowGroups.find((g) => g.metric === "foreign-buy")!.rows[0],
      ).toMatchObject({ days: 4, amount: 6000000, end: "2026-11-05" });
    } finally {
      db.close();
    }
  });
  it("resets trading-day knowledge on date rollover and does not remind on an unverified holiday", async () => {
    const db = openDatabase(":memory:");
    let now = new Date("2026-11-05T08:00:00Z");
    const notify = vi.fn();
    let fail = false;
    try {
      const s = createMarketRecordsService({
        db,
        now: () => now,
        fetchIndex: async () => [],
        fetchTradingDay: async () => {
          if (fail) throw new Error("offline");
          return true;
        },
        notify,
      });
      await s.refresh();
      fail = true;
      now = new Date("2026-11-06T08:00:00Z");
      await s.refresh();
      expect(notify).toHaveBeenCalledTimes(1);
    } finally {
      db.close();
    }
  });
  it("calculates reported-date and current index ranks and only notifies once for each confirmed direction", async () => {
    const db = openDatabase(":memory:");
    let now = new Date("2026-06-12T08:00:00Z");
    const notify = vi.fn();
    const rows = [
      {
        Date: "1150430",
        HighestIndex: "39000",
        LowestIndex: "38000",
        ClosingIndex: "38926.63",
      },
      {
        Date: "1150504",
        HighestIndex: "40755.52",
        LowestIndex: "39000",
        ClosingIndex: "40705.14",
      },
      {
        Date: "1150611",
        HighestIndex: "43149.46",
        LowestIndex: "43149.46",
        ClosingIndex: "43149.46",
      },
      {
        Date: "1150612",
        HighestIndex: "44798.88",
        LowestIndex: "40000",
        ClosingIndex: "44169.04",
      },
    ];
    try {
      const s = createMarketRecordsService({
        db,
        now: () => now,
        fetchIndex: async () => rows,
        notify,
      });
      await s.tick();
      const count = notify.mock.calls.length;
      expect(count).toBeGreaterThan(0);
      await s.tick();
      await s.refresh();
      expect(notify).toHaveBeenCalledTimes(count);
      const news = s.snapshot().news.find((n) => n.id === "index-20260612")!;
      expect(news.actual).toBe(1649.42);
      expect(news.reportedRank).toBe(2);
      expect(news.calculatedRank).toBeDefined();
      expect(
        s.snapshot().news.find((n) => n.id === "close-20260504")!.status,
      ).toBe("metric-mismatch");
      now = new Date("2026-06-13T00:00:00Z");
      await s.tick();
      expect(notify).toHaveBeenCalledTimes(count);
    } finally {
      db.close();
    }
  });
  it("preserves conflicting trust news claims, rejects future rows and coalesces simultaneous refreshes", async () => {
    const db = openDatabase(":memory:");
    let release!: (v: unknown) => void;
    const fetchIndex = vi.fn(
      () =>
        new Promise((r) => {
          release = r;
        }),
    );
    try {
      const store = createMarketRecordsStore(db);
      store.importRows(
        [
          { date: "2026-06-23", investor: "trust", net: 100000000 },
          {
            date: "2026-08-06",
            investor: "trust",
            net: 380989000000 - 100000000,
          },
        ],
        [],
        "CMoney",
        "2026-08-07",
      );
      const s = createMarketRecordsService({
        db,
        now: () => new Date("2026-08-07T08:00:00Z"),
        fetchIndex,
        notify: () => false,
      });
      const a = s.refresh(),
        b = s.refresh();
      expect(fetchIndex).toHaveBeenCalledOnce();
      release([]);
      await Promise.all([a, b]);
      expect(
        s.snapshot().news.find((n) => n.id === "trust-20260807-early")!.status,
      ).toBe("unresolved");
      expect(
        s.snapshot().news.find((n) => n.id === "trust-20260807-late")!.actual,
      ).toBe(3809.89);
      expect(() =>
        s.importFile(
          Buffer.from(
            "116年01月01日 三大法人買賣金額統計表\n自營商(自行買賣),1,0,1\n自營商(避險),1,0,1\n投信,1,0,1\n外資及陸資(不含外資自營商),1,0,1",
          ),
          "csv",
        ),
      ).toThrow("未來");
    } finally {
      db.close();
    }
  });
  it("reminds even when index lags if official calendar confirms a trading day; retries suppressed delivery", async () => {
    const db = openDatabase(":memory:");
    let deliver = false;
    const notify = vi.fn(() => deliver);
    try {
      const service = createMarketRecordsService({
        db,
        now: () => new Date("2026-11-05T08:00:00Z"),
        fetchIndex: async () => [],
        fetchTradingDay: async () => true,
        notify,
      });
      await service.refresh();
      deliver = true;
      await service.refresh();
      await service.refresh();
      expect(notify).toHaveBeenCalledTimes(2);
    } finally {
      db.close();
    }
  });
  it("does not mix unverified foreign classifications or repeat same-day record notices", () => {
    const db = openDatabase(":memory:");
    const notify = vi.fn();
    try {
      const s = createMarketRecordsService({
        db,
        now: () => new Date("2026-11-05T08:00:00Z"),
        fetchIndex: async () => [],
        notify,
      });
      const x = workbook([
        ["日期", "上市買賣超金額(百萬)"],
        ["20261103", "10"],
        ["20261104", "20"],
      ]);
      s.importFile(x, "foreign");
      const csv = (day: string, amount: string) =>
        Buffer.from(
          `115年11月${day}日 三大法人買賣金額統計表\n自營商(自行買賣),100,0,100\n自營商(避險),100,0,100\n投信,${amount},0,${amount}\n外資及陸資(不含外資自營商),100,0,100`,
        );
      s.importFile(csv("04", "100"), "csv");
      s.importFile(csv("05", "200"), "csv");
      const count = notify.mock.calls.length;
      expect(count).toBe(2);
      s.importFile(csv("05", "300"), "csv");
      expect(notify).toHaveBeenCalledTimes(count);
      expect(s.snapshot().foreignWarning).toContain("不混接");
      expect(
        s
          .snapshot()
          .flowGroups.filter((g) => g.metric.startsWith("foreign"))
          .flatMap((g) => g.rows)
          .some((r) => r.end === "2026-11-05"),
      ).toBe(false);
    } finally {
      db.close();
    }
  });
  it("does not touch live sources until refresh, reminds once after close and retains source errors", async () => {
    const db = openDatabase(":memory:");
    const notify = vi.fn();
    const fetchIndex = vi
      .fn()
      .mockResolvedValue([
        { 日期: "1151105", 最高指數: "110", 最低指數: "90", 收盤指數: "100" },
      ]);
    try {
      const service = createMarketRecordsService({
        db,
        now: () => new Date("2026-11-05T08:00:00Z"),
        fetchIndex,
        notify,
      });
      expect(fetchIndex).not.toHaveBeenCalled();
      await service.refresh();
      await service.refresh();
      expect(notify).toHaveBeenCalledTimes(1);
      expect(service.snapshot().manualRequired).toBe(true);
      fetchIndex.mockRejectedValue(new Error("offline"));
      await service.refresh();
      expect(service.snapshot().error).toContain("offline");
      expect(service.snapshot().indices.rise).toEqual([]);
    } finally {
      db.close();
    }
  });
});
