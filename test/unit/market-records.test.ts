import { describe, it, expect } from "vitest";
import {
  rankStreaks,
  rankIndex,
  compareClaim,
} from "../../src/domain/market-records";
import {
  parseTwseCsv,
  parseIndexResponse,
  parseCMoney,
  officialTradingDay,
} from "../../src/providers/market-records";
import { workbook } from "../support/market-xlsx";

describe("market-records", () => {
  it("imports synthetic CMoney dates and million-unit amounts but leaves today unconfirmed", () => {
    const data = workbook([
      ["日期", "上市買賣超金額(百萬)"],
      ["20261104", "12.34"],
      ["20261105", "-2.50"],
    ]);
    expect(parseCMoney(data, "trust", "2026-11-05").flows).toEqual([
      { date: "2026-11-04", investor: "trust", net: 12340000, confirmed: true },
      {
        date: "2026-11-05",
        investor: "trust",
        net: -2500000,
        confirmed: false,
      },
    ]);
    expect(() =>
      parseCMoney(
        workbook([
          ["日期", "上市買賣超金額(百萬)"],
          ["20261104", "bad"],
        ]),
        "trust",
        "2026-11-05",
      ),
    ).toThrow();
    expect(
      officialTradingDay(
        [{ Date: "1150101", Name: "開國紀念日", Description: "放假" }],
        "2026-11-05",
      ),
    ).toBe(true);
    expect(
      officialTradingDay(
        [{ Date: "1151105", Name: "市場無交易", Description: "" }],
        "2026-11-05",
      ),
    ).toBe(false);
  });
  it("ranks whole streaks, breaks on missing sessions and zero, and shares tie ranks", () => {
    const days = [
      "2026-01-02",
      "2026-01-05",
      "2026-01-06",
      "2026-01-07",
      "2026-01-08",
      "2026-01-09",
    ];
    const rows = days.map((date, i) => ({
      date,
      investor: "trust" as const,
      net: [10, 20, 0, 15, 15, -10][i],
    }));
    expect(
      rankStreaks(rows, days, "trust", "buy").map((r) => [
        r.start,
        r.end,
        r.amount,
        r.rank,
      ]),
    ).toEqual([
      ["2026-01-02", "2026-01-05", 30, 1],
      ["2026-01-07", "2026-01-08", 30, 1],
    ]);
    expect(
      rankStreaks(
        rows.filter((r) => r.date !== "2026-01-05"),
        days,
        "trust",
        "buy",
      ),
    ).toHaveLength(1);
  });
  it("uses previous trading close and never confirms today from a local spreadsheet", () => {
    const rows = [
      {
        date: "2026-05-01",
        high: 39000,
        low: 38000,
        close: 38926.63,
        confirmed: true,
      },
      {
        date: "2026-05-04",
        high: 40755.52,
        low: 39000,
        close: 40705.14,
        confirmed: false,
      },
    ];
    expect(rankIndex(rows, "rise", "2026-05-04")).toEqual([]);
    expect(
      rankIndex(
        rows.map((r) => ({ ...r, confirmed: true })),
        "rise",
        "2026-05-04",
      )[0],
    ).toMatchObject({ amount: 1828.89, date: "2026-05-04", rank: 1 });
    expect(
      rankIndex(
        rows.map((r) => ({ ...r, confirmed: true })),
        "rise",
        "2026-05-04",
        ["2026-05-01", "2026-05-02", "2026-05-04"],
      ),
    ).toEqual([]);
  });
  it("preserves news precision, metric conflicts, and unresolved discrepancies", () => {
    expect(
      compareClaim(
        { metric: "trust-buy", value: 3809, precision: 1 },
        "trust-buy",
        3809.89,
      ),
    ).toBe("rounded");
    expect(
      compareClaim(
        { metric: "trust-buy", value: 3813.56, precision: 0.01 },
        "trust-buy",
        3809.89,
      ),
    ).toBe("unresolved");
    expect(
      compareClaim(
        { metric: "close-rise", value: 1778.51, precision: 0.01 },
        "index-rise",
        1828.89,
      ),
    ).toBe("metric-mismatch");
  });
  it("parses quoted TWSE CSV, dates and dealer components without double counting", () => {
    const csv =
      '"115年11月04日 三大法人買賣金額統計表"\n"單位名稱","買進金額","賣出金額","買賣差額"\n"自營商(自行買賣)","100","20","80"\n"自營商(避險)","50","70","-20"\n"投信","100","0","100"\n"外資及陸資(不含外資自營商)","10","30","-20"\n"外資自營商","5","0","5"';
    expect(parseTwseCsv(csv)).toEqual([
      { date: "2026-11-04", investor: "foreign", net: -20 },
      { date: "2026-11-04", investor: "trust", net: 100 },
      { date: "2026-11-04", investor: "dealer", net: 60 },
    ]);
    expect(() => parseTwseCsv(csv.replace('"投信"', '"其他"'))).toThrow();
    expect(() =>
      parseTwseCsv(csv.replace('"100","0","100"', '"100","0","NaN"')),
    ).toThrow();
  });
  it("validates official index rows rather than treating invalid responses as no data", () => {
    expect(
      parseIndexResponse([
        {
          Date: "1151104",
          HighestIndex: "44000",
          LowestIndex: "42000",
          ClosingIndex: "43000",
        },
      ])[0],
    ).toMatchObject({ date: "2026-11-04", confirmed: true });
    expect(
      parseIndexResponse([
        {
          日期: "1151104",
          最高指數: "44,000.00",
          最低指數: "42,000.00",
          收盤指數: "43,000.00",
        },
      ])[0],
    ).toMatchObject({ date: "2026-11-04", confirmed: true });
    expect(() => parseIndexResponse("<html>blocked</html>")).toThrow();
  });
});
