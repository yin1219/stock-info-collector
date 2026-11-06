import type { Database } from "better-sqlite3";
import {
  validDate,
  type FlowDay,
  type IndexDay,
} from "../domain/market-records";
export function createMarketRecordsStore(db: Database) {
  return {
    importRows(
      flows: FlowDay[],
      indices: IndexDay[],
      source: string,
      at: string,
    ) {
      const seen = new Set<string>();
      for (const row of [...flows, ...indices]) {
        const key =
          "investor" in row ? `${row.date}:${row.investor}` : row.date;
        if (seen.has(key) || !validDate(row.date))
          throw new Error("資料日期重複或無效");
        seen.add(key);
        if (
          "net" in row
            ? !["foreign", "trust", "dealer"].includes(row.investor) ||
              !Number.isSafeInteger(row.net)
            : ![row.high, row.low, row.close].every(Number.isFinite) ||
              row.low > row.high ||
              row.close < row.low ||
              row.close > row.high
        )
          throw new Error("日資料數值無效");
      }
      db.transaction(() => {
        for (const row of [...flows, ...indices]) {
          const kind = "investor" in row ? row.investor : "index";
          const data = JSON.stringify({ ...row, source });
          const old = db
            .prepare("SELECT data FROM market_days WHERE date=? AND kind=?")
            .get(row.date, kind) as { data: string } | undefined;
          if (old && kind === "foreign") {
            const prior = JSON.parse(old.data) as FlowDay;
            const incoming = { ...row, source } as FlowDay;
            const cm =
              prior.source === "CMoney"
                ? prior
                : incoming.source === "CMoney"
                  ? incoming
                  : undefined;
            const tw =
              prior.source === "TWSE CSV"
                ? prior
                : incoming.source === "TWSE CSV"
                  ? incoming
                  : undefined;
            if (cm && tw && Number.isSafeInteger(tw.foreignDealerNet))
              db.prepare(
                "INSERT INTO market_foreign_comparisons(date,cmoney_net,twse_net,dealer_net) VALUES(?,?,?,?) ON CONFLICT(date) DO UPDATE SET cmoney_net=excluded.cmoney_net,twse_net=excluded.twse_net,dealer_net=excluded.dealer_net",
              ).run(row.date, cm.net, tw.net, tw.foreignDealerNet);
            else if (
              incoming.source === "TWSE CSV" &&
              Number.isSafeInteger(incoming.foreignDealerNet)
            )
              db.prepare(
                "UPDATE market_foreign_comparisons SET twse_net=?,dealer_net=? WHERE date=?",
              ).run(incoming.net, incoming.foreignDealerNet, row.date);
          }
          if (
            old &&
            source === "CMoney" &&
            String(JSON.parse(old.data).source).startsWith("TWSE")
          )
            continue;
          if (old && old.data !== data)
            db.prepare(
              "INSERT INTO market_revisions(date,kind,previous_data,replacement_data,imported_at) VALUES(?,?,?,?,?)",
            ).run(row.date, kind, old.data, data, at);
          db.prepare(
            "INSERT INTO market_days(date,kind,data,imported_at) VALUES(?,?,?,?) ON CONFLICT(date,kind) DO UPDATE SET data=excluded.data,imported_at=excluded.imported_at",
          ).run(row.date, kind, data, at);
        }
      })();
    },
    flows: () =>
      (
        db
          .prepare(
            "SELECT data FROM market_days WHERE kind!='index' ORDER BY date",
          )
          .all() as { data: string }[]
      ).map((r) => JSON.parse(r.data) as FlowDay),
    indices: () =>
      (
        db
          .prepare(
            "SELECT data FROM market_days WHERE kind='index' ORDER BY date",
          )
          .all() as { data: string }[]
      ).map((r) => JSON.parse(r.data) as IndexDay),
    revisions: () =>
      db.prepare("SELECT * FROM market_revisions ORDER BY id").all(),
    claimNotice: (key: string) =>
      db.prepare("INSERT OR IGNORE INTO market_notices(key) VALUES(?)").run(key)
        .changes === 1,
    hasNotice: (key: string) =>
      Boolean(
        db.prepare("SELECT key FROM market_notices WHERE key=?").get(key),
      ),
    foreignBasis(): "including" | "excluding" | undefined {
      const rows = db
        .prepare("SELECT * FROM market_foreign_comparisons")
        .all() as {
        cmoney_net: number;
        twse_net: number;
        dealer_net: number;
      }[];
      if (rows.length < 3) return undefined;
      const excluding = rows.every(
        (r) => Math.abs(r.cmoney_net - r.twse_net) <= 10000,
      );
      const including = rows.every(
        (r) => Math.abs(r.cmoney_net - r.twse_net - r.dealer_net) <= 10000,
      );
      return including !== excluding
        ? including
          ? "including"
          : "excluding"
        : undefined;
    },
  };
}
