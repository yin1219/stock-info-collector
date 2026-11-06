import type { Database } from "better-sqlite3";
import {
  compareClaim,
  investorLabels,
  rankIndex,
  rankStreaks,
  type Investor,
} from "../domain/market-records";
import { marketNews } from "../data/market-news";
import {
  parseCMoney,
  parseIndexResponse,
  parseTwseCsv,
} from "../providers/market-records";
import { createMarketRecordsStore } from "../repositories/market-records";
export type MarketImportKind = Investor | "index" | "csv";
export function createMarketRecordsService(deps: {
  db: Database;
  now: () => Date;
  fetchIndex: () => Promise<unknown>;
  fetchTradingDay?: (date: string) => Promise<boolean | undefined>;
  notify: (title: string, body: string) => void | boolean;
  automaticUpdatesEnabled?: boolean;
}) {
  const store = createMarketRecordsStore(deps.db);
  let error: string | null = null;
  let indexCheckedAt: string | null = null;
  let active: Promise<unknown> | undefined;
  let lastAttempt = "";
  let tradingDay: boolean | undefined;
  const date = () =>
    new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Taipei",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).format(deps.now());
  function snapshot() {
    const today = date(),
      flows = store.flows(),
      indices = store.indices();
    const calendar = [
      ...new Set([...indices.map((r) => r.date), ...flows.map((r) => r.date)]),
    ]
      .filter((d) => d <= today)
      .sort();
    const savedIndexDates = new Set(indices.map((row) => row.date));
    const indexUpdate = {
      automatic: deps.automaticUpdatesEnabled !== false,
      latestOfficialDate:
        indices
          .filter((row) => row.source === "TWSE OpenAPI" && row.confirmed)
          .at(-1)?.date ?? null,
      checkedAt: indexCheckedAt,
      knownMissingDates: calendar.filter(
        (day) => day < today && !savedIndexDates.has(day),
      ),
    };
    const mapping = store.foreignBasis();
    const foreignBasis = flows.some(
      (r) => r.investor === "foreign" && r.source === "CMoney",
    )
      ? "CMoney"
      : "TWSE CSV";
    const confirmed = flows
      .filter(
        (r) =>
          r.confirmed !== false &&
          r.date <= today &&
          (r.investor !== "foreign" ||
            r.source === foreignBasis ||
            Boolean(mapping)),
      )
      .filter(
        (r) =>
          r.investor !== "foreign" ||
          r.source !== "TWSE CSV" ||
          mapping !== "including" ||
          r.foreignDealerNet !== undefined,
      )
      .map((r) =>
        r.investor === "foreign" &&
        r.source === "TWSE CSV" &&
        mapping === "including"
          ? { ...r, net: r.net + r.foreignDealerNet! }
          : r,
      );
    const allFlowGroups = (
      ["foreign", "trust", "dealer"] as Investor[]
    ).flatMap((investor) =>
      (["buy", "sell"] as const).map((direction) => ({
        metric: `${investor}-${direction}`,
        label: `${investorLabels[investor]}連續${direction === "buy" ? "買" : "賣"}超`,
        source:
          investor === "foreign"
            ? mapping
              ? `CMoney／TWSE 多日核對：${mapping === "including" ? "含" : "不含"}外資自營商`
              : foreignBasis
            : "CMoney／TWSE（相同法人分類）",
        rows: rankStreaks(confirmed, calendar, investor, direction),
      })),
    );
    const flowGroups = allFlowGroups.map((group) => ({
      ...group,
      rows: group.rows.filter((row) => row.rank <= 3),
    }));
    const latestFlows = (["foreign", "trust", "dealer"] as Investor[]).map(
      (investor) => {
        const dataDate =
          confirmed
            .filter((row) => row.investor === investor)
            .map((row) => row.date)
            .sort()
            .at(-1) ?? null;
        const group = allFlowGroups.find(
          (group) =>
            group.metric.startsWith(`${investor}-`) &&
            group.rows.some((row) => row.end === dataDate),
        );
        return {
          investor,
          dataDate,
          direction: group
            ? group.metric.endsWith("-buy")
              ? ("buy" as const)
              : ("sell" as const)
            : null,
          streak: group?.rows.find((row) => row.end === dataDate) ?? null,
        };
      },
    );
    const indexRanks = {
      rise: rankIndex(indices, "rise", today, calendar),
      fall: rankIndex(indices, "fall", today, calendar),
    };
    const news = marketNews.map((claim) => {
      const [investor, direction] = claim.metric.split("-");
      let actual: number | undefined,
        currentRank: number | undefined,
        calculatedRank: number | undefined;
      if (investor === "index" || investor === "close") {
        const list = indexRanks[direction === "fall" ? "fall" : "rise"];
        const row = list.find((r) => r.date === claim.end);
        actual = row?.amount;
        currentRank = row?.rank;
        calculatedRank = rankIndex(
          indices.filter((r) => r.date <= claim.end),
          direction === "fall" ? "fall" : "rise",
          claim.end,
          calendar.filter((d) => d <= claim.end),
        ).find((r) => r.date === claim.end)?.rank;
      } else {
        const historical = rankStreaks(
          confirmed.filter((r) => r.date <= claim.end),
          calendar.filter((d) => d <= claim.end),
          investor as Investor,
          direction as "buy" | "sell",
        );
        const row = historical.find(
          (r) => r.start === claim.start && r.end === claim.end,
        );
        actual = row ? row.amount / 1e8 : undefined;
        calculatedRank = row?.rank;
        currentRank = rankStreaks(
          confirmed,
          calendar,
          investor as Investor,
          direction as "buy" | "sell",
        ).find((r) => r.start === claim.start)?.rank;
      }
      const status = compareClaim(
        claim,
        investor === "close" ? "index-rise" : claim.metric,
        actual,
      );
      return {
        ...claim,
        actual,
        currentRank,
        calculatedRank,
        status: claim.note && status !== "missing-data" ? "unresolved" : status,
      };
    });
    return {
      today,
      coverage: { start: calendar[0] ?? null, end: calendar.at(-1) ?? null },
      manualRequired: !(["foreign", "trust", "dealer"] as Investor[]).every(
        (i) =>
          flows.some(
            (r) =>
              r.date === today && r.investor === i && r.confirmed !== false,
          ),
      ),
      todayConfirmed: indices.some((r) => r.date === today && r.confirmed),
      indexUpdate,
      error,
      flowGroups,
      latestFlows,
      indices: indexRanks,
      news,
      sourceStatus: "FinMind 數值及免費授權尚未通過確認；請手動匯入 TWSE CSV",
      foreignWarning:
        !mapping &&
        flows.some((r) => r.investor === "foreign" && r.source !== foreignBasis)
          ? "CMoney 外資與 TWSE 排除外資自營商的分類尚未核對；外資歷史排名只使用 CMoney，不混接 CSV。請匯入至少三個與 CMoney 重疊日期的 TWSE CSV（含外資自營商列），系統核對且可區分兩種口徑後才銜接。"
          : null,
      revisions: store.revisions().length,
    };
  }
  function notice(key: string, title: string, body: string) {
    if (store.hasNotice(key)) return;
    const result = deps.notify(title, body);
    if (result !== false) store.claimNotice(key);
  }
  function notifications() {
    const s = snapshot();
    for (const group of s.flowGroups)
      for (const row of group.rows)
        if (row.end === s.today)
          notice(
            `record:${s.today}:${group.metric}`,
            "盤面紀錄",
            `${group.label}累計金額進入2020年以來第${row.rank}名`,
          );
    for (const direction of ["rise", "fall"] as const) {
      const row = s.indices[direction].find(
        (r) => r.date === s.today && r.rank <= 3,
      );
      if (row)
        notice(
          `record:${s.today}:index-${direction}`,
          "盤面紀錄",
          `今日盤中${direction === "rise" ? "漲" : "跌"}${row.amount}點，2020年以來第${row.rank}名`,
        );
    }
  }
  async function refresh() {
    if (active) return active;
    active = (async () => {
      indexCheckedAt = deps.now().toISOString();
      try {
        const rows = parseIndexResponse(await deps.fetchIndex());
        if (rows.some((r) => r.date > date()))
          throw new Error("官方資料包含未來日期");
        store.importRows([], rows, "TWSE OpenAPI", deps.now().toISOString());
        error = null;
      } catch (e) {
        error = e instanceof Error ? e.message : String(e);
      }
      notifications();
      const s = snapshot();
      const local = new Intl.DateTimeFormat("en-GB", {
        timeZone: "Asia/Taipei",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      }).format(deps.now());
      tradingDay = undefined;
      if (deps.fetchTradingDay)
        try {
          tradingDay = await deps.fetchTradingDay(s.today);
        } catch (e) {
          error = `${error ?? ""} 開休市日曆無法確認：${String(e)}`;
        }
      if (
        (s.todayConfirmed || tradingDay === true) &&
        s.manualRequired &&
        local >= "15:30"
      )
        notice(
          `import:${s.today}`,
          "法人日資料待更新",
          "請到 TWSE 下載當日 CSV，回盤面紀錄頁匯入。",
        );
      return snapshot();
    })().finally(() => {
      active = undefined;
    });
    return active;
  }
  return {
    snapshot,
    refresh,
    async tick() {
      if (deps.automaticUpdatesEnabled === false) return;
      const local = new Intl.DateTimeFormat("en-GB", {
        timeZone: "Asia/Taipei",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      }).format(deps.now());
      const slot = `${date()}:${local.slice(0, 2)}`;
      if (local < "15:30" || lastAttempt === slot) return;
      lastAttempt = slot;
      return refresh();
    },
    importFile(buffer: Buffer, kind: MarketImportKind) {
      const parsed =
        kind === "csv"
          ? {
              flows: parseTwseCsv(
                new TextDecoder("big5").decode(buffer).includes("三大法人")
                  ? new TextDecoder("big5").decode(buffer)
                  : buffer.toString("utf8"),
                true,
              ).map((r) => ({ ...r, confirmed: true })),
              indices: [],
            }
          : parseCMoney(buffer, kind, date());
      if ([...parsed.flows, ...parsed.indices].some((r) => r.date > date()))
        throw new Error("不能匯入未來日期");
      store.importRows(
        parsed.flows,
        parsed.indices,
        kind === "csv" ? "TWSE CSV" : "CMoney",
        deps.now().toISOString(),
      );
      notifications();
      return {
        status: "imported" as const,
        count: parsed.flows.length + parsed.indices.length,
        snapshot: snapshot(),
      };
    },
  };
}
export type MarketSnapshot = ReturnType<
  ReturnType<typeof createMarketRecordsService>["snapshot"]
>;
