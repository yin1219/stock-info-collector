import { inflateRawSync } from "node:zlib";
import {
  validDate,
  type FlowDay,
  type IndexDay,
  type Investor,
} from "../domain/market-records";

function dateValue(value: string): string {
  const text = value.replace(/\D/g, "");
  let y: number, m: string, d: string;
  if (text.length === 7) {
    y = Number(text.slice(0, 3)) + 1911;
    m = text.slice(3, 5);
    d = text.slice(5);
  } else if (text.length === 8) {
    y = Number(text.slice(0, 4));
    m = text.slice(4, 6);
    d = text.slice(6);
  } else throw new Error("資料日期格式無效");
  const result = `${y}-${m}-${d}`;
  if (!validDate(result)) throw new Error("資料日期無效或早於2020年");
  return result;
}
function numeric(value: unknown): number {
  if (typeof value !== "string" && typeof value !== "number")
    throw new Error("資料數值缺漏");
  const clean = String(value).replaceAll(",", "").trim();
  if (!/^-?\d+(\.\d+)?$/.test(clean)) throw new Error("資料數值無效");
  const number = Number(clean);
  if (!Number.isFinite(number)) throw new Error("資料數值無效");
  return number;
}
export function csvRows(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [],
    cell = "",
    quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"') {
      if (quoted && text[i + 1] === '"') {
        cell += '"';
        i++;
      } else quoted = !quoted;
    } else if (c === "," && !quoted) {
      row.push(cell.trim());
      cell = "";
    } else if (c === "\n" && !quoted) {
      row.push(cell.trim());
      rows.push(row);
      row = [];
      cell = "";
    } else if (c !== "\r") cell += c;
  }
  if (quoted) throw new Error("CSV引號不完整");
  if (cell || row.length) {
    row.push(cell.trim());
    rows.push(row);
  }
  return rows;
}
export function parseTwseCsv(text: string, components = false): FlowDay[] {
  const match = text.match(/(\d{3})年(\d{1,2})月(\d{1,2})日/);
  if (!match) throw new Error("找不到TWSE日報日期");
  const date = dateValue(
    `${match[1]}${match[2].padStart(2, "0")}${match[3].padStart(2, "0")}`,
  );
  const rows = csvRows(text);
  const get = (name: string) => {
    const row = rows.find((r) => r[0] === name);
    if (!row || row.length < 4) throw new Error(`缺少分類：${name}`);
    const buy = numeric(row[1]),
      sell = numeric(row[2]),
      net = numeric(row[3]);
    if (buy < 0 || sell < 0 || buy - sell !== net || !Number.isSafeInteger(net))
      throw new Error("買賣金額不一致");
    return net;
  };
  const foreignDealerNet =
    components && rows.some((r) => r[0] === "外資自營商")
      ? get("外資自營商")
      : undefined;
  return [
    {
      date,
      investor: "foreign",
      net: get("外資及陸資(不含外資自營商)"),
      ...(foreignDealerNet === undefined ? {} : { foreignDealerNet }),
    },
    { date, investor: "trust", net: get("投信") },
    {
      date,
      investor: "dealer",
      net: get("自營商(自行買賣)") + get("自營商(避險)"),
    },
  ];
}
export function parseIndexResponse(value: unknown): IndexDay[] {
  if (!Array.isArray(value)) throw new Error("官方指數回應格式無效");
  return value.map((row) => {
    const r = row as Record<string, unknown>;
    const result = {
      date: dateValue(String(r.Date ?? r.日期)),
      high: numeric(r.HighestIndex ?? r.最高指數),
      low: numeric(r.LowestIndex ?? r.最低指數),
      close: numeric(r.ClosingIndex ?? r.收盤指數),
      confirmed: true,
      source: "TWSE OpenAPI",
    };
    if (
      result.low > result.high ||
      result.close < result.low ||
      result.close > result.high
    )
      throw new Error("指數高低值無效");
    return result;
  });
}
export function officialTradingDay(
  value: unknown,
  date: string,
): boolean | undefined {
  if (!Array.isArray(value) || !value.length)
    throw new Error("官方開休市日期格式無效");
  const rows = value.map((r) => ({
    date: dateValue(String(r.Date)),
    name: String(r.Name),
    description: String(r.Description),
  }));
  if (!rows.some((r) => r.date.slice(0, 4) === date.slice(0, 4)))
    return undefined;
  const row = rows.find((r) => r.date === date);
  if (row)
    return /開始交易|最後交易/.test(row.name) && !/無交易/.test(row.name);
  return ![0, 6].includes(new Date(`${date}T12:00:00+08:00`).getUTCDay());
}
// Read only the bounded OOXML entries needed by a single-sheet CMoney export.
function unzip(buffer: Buffer): Map<string, string> {
  if (buffer.length > 20_000_000) throw new Error("Excel檔案超過20MB");
  let eocd = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 65557); i--)
    if (buffer.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  if (eocd < 0) throw new Error("Excel壓縮格式無效");
  const files = new Map<string, string>();
  let offset = buffer.readUInt32LE(eocd + 16);
  for (let n = 0; n < buffer.readUInt16LE(eocd + 10); n++) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50)
      throw new Error("Excel目錄無效");
    const method = buffer.readUInt16LE(offset + 10),
      size = buffer.readUInt32LE(offset + 20),
      plain = buffer.readUInt32LE(offset + 24),
      len = buffer.readUInt16LE(offset + 28),
      extra = buffer.readUInt16LE(offset + 30),
      comment = buffer.readUInt16LE(offset + 32),
      local = buffer.readUInt32LE(offset + 42);
    const name = buffer
      .subarray(offset + 46, offset + 46 + len)
      .toString("utf8");
    offset += 46 + len + extra + comment;
    if (!["xl/sharedStrings.xml", "xl/worksheets/sheet1.xml"].includes(name))
      continue;
    if (plain > 12_000_000 || local + 30 > buffer.length)
      throw new Error("Excel內容過大或無效");
    const start =
      local +
      30 +
      buffer.readUInt16LE(local + 26) +
      buffer.readUInt16LE(local + 28);
    const compressed = buffer.subarray(start, start + size);
    const data =
      method === 0
        ? compressed
        : method === 8
          ? inflateRawSync(compressed, { maxOutputLength: 12_000_000 })
          : undefined;
    if (!data) throw new Error("Excel壓縮方式不支援");
    files.set(name, data.toString("utf8"));
  }
  return files;
}
const unescapeXml = (s: string) =>
  s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
export function parseCMoney(
  buffer: Buffer,
  investor: Investor | "index",
  today: string,
): { flows: FlowDay[]; indices: IndexDay[] } {
  const files = unzip(buffer);
  const strings = [
    ...(files.get("xl/sharedStrings.xml") ?? "").matchAll(
      /<si\b[^>]*>([\s\S]*?)<\/si>/g,
    ),
  ].map((m) =>
    [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
      .map((t) => unescapeXml(t[1]))
      .join(""),
  );
  const xml = files.get("xl/worksheets/sheet1.xml");
  if (!xml) throw new Error("Excel缺少第一工作表");
  const rows = [...xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)].map((m) => {
    const row: string[] = [];
    for (const cell of m[1].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const column = cell[1].match(/r="([A-Z]+)\d+"/)?.[1];
      if (!column) continue;
      let index = 0;
      for (const c of column) index = index * 26 + c.charCodeAt(0) - 64;
      const raw =
        cell[2].match(/<v>([\s\S]*?)<\/v>/)?.[1] ??
        cell[2].match(/<t[^>]*>([\s\S]*?)<\/t>/)?.[1] ??
        "";
      row[index - 1] = cell[1].includes('t="s"')
        ? strings[Number(raw)]
        : unescapeXml(raw);
    }
    return row;
  });
  if (
    rows[0]?.[0] !== "日期" ||
    (investor === "index"
      ? rows[0]?.[2] !== "最高價"
      : rows[0]?.[1] !== "上市買賣超金額(百萬)")
  )
    throw new Error("CMoney欄位或單位不符");
  const flows: FlowDay[] = [],
    indices: IndexDay[] = [];
  for (const row of rows.slice(1)) {
    if (!row[0]) continue;
    const date = /^\d{8}$/.test(row[0])
      ? dateValue(row[0])
      : /^\d{4}[-/]\d/.test(row[0])
        ? dateValue(row[0])
        : new Date((numeric(row[0]) - 25569) * 86400000)
            .toISOString()
            .slice(0, 10);
    if (!validDate(date) || date > today) throw new Error("Excel日期無效");
    if (investor === "index") {
      const bar = {
        date,
        high: numeric(row[2]),
        low: numeric(row[3]),
        close: numeric(row[4]),
        confirmed: date < today,
        source: "CMoney",
      };
      if (bar.low > bar.high || bar.close < bar.low || bar.close > bar.high)
        throw new Error("指數高低值無效");
      indices.push(bar);
    } else
      flows.push({
        date,
        investor,
        net: Math.round(numeric(row[1]) * 1_000_000),
        confirmed: date < today,
      });
  }
  if (!flows.length && !indices.length) throw new Error("Excel沒有日資料");
  return { flows, indices };
}
