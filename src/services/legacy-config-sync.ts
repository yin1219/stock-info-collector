import { createHash, randomUUID } from 'node:crypto';
import { lstat, open, readFile, rename, unlink } from 'node:fs/promises';
import path from 'node:path';

function fingerprint(contents: string): string {
  return createHash('sha256').update(contents, 'utf8').digest('hex');
}

function parseConfig(contents: string): Record<string, unknown> & { StockNumbers: string } {
  let parsed: unknown;
  try {
    parsed = JSON.parse(contents);
  } catch {
    throw new Error('既有 v1 設定不是有效的 JSON');
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)
    || typeof (parsed as Record<string, unknown>).StockNumbers !== 'string') {
    throw new Error('既有 v1 設定缺少 StockNumbers 字串');
  }
  return parsed as Record<string, unknown> & { StockNumbers: string };
}

async function readRegularConfig(configPath: string): Promise<string> {
  const details = await lstat(configPath);
  if (!details.isFile() || details.isSymbolicLink()) throw new Error('v1 設定不是一般檔案，未套用關注股號');
  return readFile(configPath, 'utf8');
}

export async function inspectLegacyConfig(configPath: string): Promise<{ hash: string; stockNumbers: string[] }> {
  const contents = await readRegularConfig(configPath);
  const config = parseConfig(contents);
  return {
    hash: fingerprint(contents),
    stockNumbers: config.StockNumbers.split(',').map((code) => code.trim()).filter(Boolean),
  };
}

async function writeNewFile(filePath: string, contents: string): Promise<void> {
  const file = await open(filePath, 'wx', 0o600);
  try {
    await file.writeFile(contents, 'utf8');
    await file.sync();
  } finally {
    await file.close();
  }
}

export async function syncLegacyStockNumbers(input: {
  configPath: string;
  reviewedHash: string;
  activeStockCodes: readonly string[];
  replaceFile?: typeof rename;
}): Promise<{ status: 'applied' | 'unchanged'; stockNumbers: string[]; backupPath?: string }> {
  const codes = [...new Set(input.activeStockCodes.map((code) => code.trim()))];
  if (codes.length === 0) throw new Error('v1 每日工作至少一間關注公司須保持啟用；未將空股號寫入舊設定');
  if (codes.some((code) => !/^\d{4,6}$/.test(code))) throw new Error('股票代號格式無效，未套用到 v1');

  const original = await readRegularConfig(input.configPath);
  if (fingerprint(original) !== input.reviewedHash) throw new Error('v1 設定已被外部修改，請重新核對後再套用');
  const config = parseConfig(original);
  const stockNumbers = codes.join(',');
  if (config.StockNumbers === stockNumbers) return { status: 'unchanged', stockNumbers: codes };

  const directory = path.dirname(input.configPath);
  const nonce = randomUUID();
  const backupPath = path.join(directory, `${path.basename(input.configPath)}.pre-v2-${nonce}.backup`);
  const temporaryPath = path.join(directory, `.${path.basename(input.configPath)}.${nonce}.tmp`);
  try {
    await writeNewFile(backupPath, original);
    await writeNewFile(temporaryPath, `${JSON.stringify({ ...config, StockNumbers: stockNumbers }, null, 2)}\n`);
    if (fingerprint(await readRegularConfig(input.configPath)) !== input.reviewedHash) {
      throw new Error('v1 設定已被外部修改，請重新核對後再套用');
    }
    await (input.replaceFile ?? rename)(temporaryPath, input.configPath);
    return { status: 'applied', stockNumbers: codes, backupPath };
  } finally {
    await unlink(temporaryPath).catch(() => undefined);
  }
}
