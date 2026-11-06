export type Investor = "foreign" | "trust" | "dealer";
export interface FlowDay {
  date: string;
  investor: Investor;
  net: number;
  source?: string;
  confirmed?: boolean;
  foreignDealerNet?: number;
}
export interface IndexDay {
  date: string;
  high: number;
  low: number;
  close: number;
  confirmed: boolean;
  source?: string;
}
export interface Streak {
  start: string;
  end: string;
  days: number;
  amount: number;
  rank: number;
}
export const investorLabels: Record<Investor, string> = {
  foreign: "外資（依來源分類）",
  trust: "投信",
  dealer: "自營商",
};
export function validDate(date: string): boolean {
  return (
    /^20\d{2}-\d{2}-\d{2}$/.test(date) &&
    date >= "2020-01-01" &&
    !Number.isNaN(Date.parse(date)) &&
    new Date(date).toISOString().slice(0, 10) === date
  );
}
export function ranked<T extends { amount: number }>(
  rows: T[],
): Array<T & { rank: number }> {
  const sorted = rows.sort((a, b) => b.amount - a.amount);
  let rank = 0;
  return sorted.map((row, i) => {
    if (i === 0 || row.amount !== sorted[i - 1].amount) rank = i + 1;
    return { ...row, rank };
  });
}
export function rankStreaks(
  rows: FlowDay[],
  calendar: string[],
  investor: Investor,
  direction: "buy" | "sell",
): Streak[] {
  const values = new Map(
    rows.filter((r) => r.investor === investor).map((r) => [r.date, r.net]),
  );
  const result: Omit<Streak, "rank">[] = [];
  let active: Omit<Streak, "rank"> | undefined;
  const finish = () => {
    if (active && active.days >= 2) result.push(active);
    active = undefined;
  };
  for (const date of [...new Set(calendar)].filter(validDate).sort()) {
    const net = values.get(date);
    const amount = net === undefined ? 0 : direction === "buy" ? net : -net;
    if (amount <= 0) {
      finish();
      continue;
    }
    active = active
      ? {
          ...active,
          end: date,
          days: active.days + 1,
          amount: active.amount + amount,
        }
      : { start: date, end: date, days: 1, amount };
  }
  finish();
  return ranked(result);
}
export function rankIndex(
  rows: IndexDay[],
  direction: "rise" | "fall",
  today: string,
  calendar: string[] = rows.map((r) => r.date),
) {
  const sorted = [...rows].sort((a, b) => a.date.localeCompare(b.date));
  const sessions = [...new Set(calendar)].sort();
  return ranked(
    sorted.flatMap((row, i) => {
      if (
        i === 0 ||
        !row.confirmed ||
        row.date > today ||
        !sorted[i - 1].confirmed
      )
        return [];
      if (sessions[sessions.indexOf(row.date) - 1] !== sorted[i - 1].date)
        return [];
      const amount =
        Math.round(
          (direction === "rise"
            ? row.high - sorted[i - 1].close
            : sorted[i - 1].close - row.low) * 100,
        ) / 100;
      return amount > 0 ? [{ date: row.date, amount }] : [];
    }),
  );
}
export interface NewsClaim {
  id: string;
  url: string;
  publisher: string;
  publishedAt: string;
  updatedAt?: string;
  observedAt?: string;
  metric: string;
  start: string;
  end: string;
  days?: number;
  value: number;
  precision: number;
  reportedRank?: number;
  scope: string;
  note?: string;
}
export type MarketImportKind = Investor | "index" | "csv";
export interface MarketSnapshot {
  indexUpdate?: {
    automatic: boolean;
    latestOfficialDate: string | null;
    checkedAt: string | null;
    knownMissingDates: string[];
  };
  latestFlows?: Array<{
    investor: Investor;
    dataDate: string | null;
    direction: "buy" | "sell" | null;
    streak: Streak | null;
  }>;
  today: string;
  coverage: { start: string | null; end: string | null };
  manualRequired: boolean;
  todayConfirmed: boolean;
  error: string | null;
  flowGroups: Array<{
    metric: string;
    label: string;
    source: string;
    rows: Streak[];
  }>;
  indices: Record<
    "rise" | "fall",
    Array<{ date: string; amount: number; rank: number }>
  >;
  news: Array<
    NewsClaim & {
      actual?: number;
      currentRank?: number;
      calculatedRank?: number;
      status: string;
    }
  >;
  sourceStatus: string;
  foreignWarning: string | null;
  revisions: number;
}
export function compareClaim(
  claim: Pick<NewsClaim, "metric" | "value" | "precision">,
  metric: string,
  actual: number | undefined,
) {
  if (claim.metric !== metric) return "metric-mismatch";
  if (actual === undefined) return "missing-data";
  if (Math.abs(claim.value - actual) < 0.000001) return "matched";
  return Math.abs(claim.value - actual) <= claim.precision
    ? "rounded"
    : "unresolved";
}
