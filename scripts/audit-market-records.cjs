// Opt-in local audit. Reads the user's exports, never connects to external services.
const fs = require("node:fs");
const path = require("node:path");
const { createJiti } = require("jiti");
const load = createJiti(__filename);
const { parseCMoney } = load("../src/providers/market-records.ts");
const { rankStreaks, rankIndex } = load("../src/domain/market-records.ts");
const { marketNews } = load("../src/data/market-news.ts");
const directory = path.resolve(process.argv[2] || "data");
const today =
  process.argv[3] ||
  new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Taipei",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date());
const flows = [],
  indices = [];
for (const file of fs
  .readdirSync(directory)
  .filter((name) => name.endsWith(".xlsx"))) {
  const kind = file.includes("自營商")
    ? "dealer"
    : file.includes("投信")
      ? "trust"
      : file.includes("外資")
        ? "foreign"
        : file.includes("加權指數")
          ? "index"
          : null;
  if (!kind) continue;
  const parsed = parseCMoney(
    fs.readFileSync(path.join(directory, file)),
    kind,
    today,
  );
  flows.push(...parsed.flows);
  indices.push(...parsed.indices);
}
const calendar = [...new Set(indices.map((row) => row.date))].sort();
const report = {
  asOf: today,
  coverage: [calendar[0], calendar.at(-1)],
  dailyCounts: { flows: flows.length, indices: indices.length },
  flows: [],
  indices: {},
  news: marketNews.length,
};
for (const investor of ["foreign", "trust", "dealer"])
  for (const direction of ["buy", "sell"]) {
    const rows = rankStreaks(
      flows.filter((row) => row.confirmed),
      calendar,
      investor,
      direction,
    ).filter((row) => row.rank <= 3);
    report.flows.push({
      investor,
      direction,
      rows: rows.map((row) => ({
        ...row,
        amountYi: row.amount / 1e8,
        evidence: marketNews
          .filter(
            (claim) =>
              claim.metric === `${investor}-${direction}` &&
              claim.start === row.start,
          )
          .map((claim) => claim.url),
      })),
    });
  }
for (const direction of ["rise", "fall"])
  report.indices[direction] = rankIndex(
    indices,
    direction,
    today,
    calendar,
  ).filter((row) => row.rank <= 3);
console.log(JSON.stringify(report, null, 2));
