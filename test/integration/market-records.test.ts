import { describe, it, expect } from "vitest";
import { openDatabase } from "../../src/repositories/migrations";
import { createMarketRecordsStore } from "../../src/repositories/market-records";
import { createTestPaths } from "../support";
describe("market-records SQLite", () => {
  it("requires three distinguishing overlaps before approving a foreign classification", () => {
    const db = openDatabase(":memory:");
    try {
      const s = createMarketRecordsStore(db);
      for (const date of ["2026-11-02", "2026-11-03", "2026-11-04"]) {
        s.importRows(
          [{ date, investor: "foreign", net: 1500000 }],
          [],
          "CMoney",
          "now",
        );
        s.importRows(
          [
            {
              date,
              investor: "foreign",
              net: 1000000,
              foreignDealerNet: 500000,
            },
          ],
          [],
          "TWSE CSV",
          "now",
        );
      }
      expect(s.foreignBasis()).toBe("including");
    } finally {
      db.close();
    }
  });
  it("retains notices and market data across a real database reopen", async () => {
    const paths = await createTestPaths();
    try {
      const first = openDatabase(paths.database);
      const a = createMarketRecordsStore(first);
      a.importRows(
        [{ date: "2026-11-04", investor: "trust", net: 100 }],
        [],
        "TWSE CSV",
        "2026-11-05",
      );
      a.claimNotice("record:2026-11-04:trust-buy");
      first.close();
      const second = openDatabase(paths.database);
      try {
        const b = createMarketRecordsStore(second);
        expect(b.flows()).toHaveLength(1);
        expect(b.claimNotice("record:2026-11-04:trust-buy")).toBe(false);
      } finally {
        second.close();
      }
    } finally {
      await paths.cleanup();
    }
  });
  it("imports atomically, retains revisions, and persists notification deduplication", () => {
    const db = openDatabase(":memory:");
    try {
      const store = createMarketRecordsStore(db);
      store.importRows(
        [{ date: "2026-11-04", investor: "trust", net: 100 }],
        [],
        "TWSE CSV",
        "2026-11-05T00:00:00Z",
      );
      store.importRows(
        [{ date: "2026-11-04", investor: "trust", net: 200 }],
        [],
        "TWSE CSV",
        "2026-11-05T01:00:00Z",
      );
      expect(store.flows()[0].net).toBe(200);
      store.importRows(
        [{ date: "2026-11-04", investor: "trust", net: 50, confirmed: false }],
        [],
        "CMoney",
        "2026-11-05T02:00:00Z",
      );
      expect(store.flows()[0].net).toBe(200);
      expect(store.revisions()).toHaveLength(1);
      expect(() =>
        store.importRows(
          [{ date: "bad", investor: "trust", net: 300 }],
          [],
          "bad",
          "now",
        ),
      ).toThrow();
      expect(store.flows()).toHaveLength(1);
      expect(store.claimNotice("record:2026-11-04:trust-buy")).toBe(true);
      expect(
        createMarketRecordsStore(db).claimNotice("record:2026-11-04:trust-buy"),
      ).toBe(false);
    } finally {
      db.close();
    }
  });
});
