import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { inspectLegacyConfig, syncLegacyStockNumbers } from '../../src/services/legacy-config-sync';

const tempDirectories: string[] = [];

async function fixture() {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'reporter-v1-config-'));
  tempDirectories.push(directory);
  const configPath = path.join(directory, 'default.json');
  const original = '{"Debug":false,"LogLevel":"info","StockNumbers":"2330, 2454","Other":{"keep":true}}\n';
  await writeFile(configPath, original, 'utf8');
  return { configPath, original };
}

afterEach(async () => {
  for (const directory of tempDirectories.splice(0)) {
    if (path.dirname(directory) !== os.tmpdir() || !path.basename(directory).startsWith('reporter-v1-config-')) {
      throw new Error('拒絕清理非測試暫存目錄');
    }
    await rm(directory, { recursive: true, force: true });
  }
});

describe('watchlist-management / v1 StockNumbers sync / safe coexistence', () => {
  it('changes only active stock codes and leaves a recoverable original backup', async () => {
    const { configPath, original } = await fixture();
    const review = await inspectLegacyConfig(configPath);
    expect(review.stockNumbers).toEqual(['2330', '2454']);

    const result = await syncLegacyStockNumbers({ configPath, reviewedHash: review.hash, activeStockCodes: ['2454', '6488', '2454'] });
    expect(result.status).toBe('applied');
    expect(result.stockNumbers).toEqual(['2454', '6488']);
    expect(JSON.parse(await readFile(configPath, 'utf8'))).toEqual({
      Debug: false, LogLevel: 'info', StockNumbers: '2454,6488', Other: { keep: true },
    });
    expect(await readFile(result.backupPath, 'utf8')).toBe(original);
  });

  it('rejects an external edit after review without overwriting the config', async () => {
    const { configPath } = await fixture();
    const review = await inspectLegacyConfig(configPath);
    const external = '{"Debug":true,"StockNumbers":"1101","Other":"external"}\n';
    await writeFile(configPath, external, 'utf8');

    await expect(syncLegacyStockNumbers({ configPath, reviewedHash: review.hash, activeStockCodes: ['2330'] }))
      .rejects.toThrow(/外部修改/);
    expect(await readFile(configPath, 'utf8')).toBe(external);
  });

  it('does not claim success or replace the original when atomic replacement fails', async () => {
    const { configPath, original } = await fixture();
    const review = await inspectLegacyConfig(configPath);
    await expect(syncLegacyStockNumbers({
      configPath, reviewedHash: review.hash, activeStockCodes: ['6488'],
      replaceFile: async () => { throw new Error('access denied'); },
    })).rejects.toThrow(/access denied/);
    expect(await readFile(configPath, 'utf8')).toBe(original);
  });

  it('rejects malformed JSON, missing StockNumbers, and invalid active codes', async () => {
    const { configPath } = await fixture();
    await writeFile(configPath, '{not-json', 'utf8');
    await expect(inspectLegacyConfig(configPath)).rejects.toThrow(/JSON/);
    await writeFile(configPath, '{"Debug":true}', 'utf8');
    await expect(inspectLegacyConfig(configPath)).rejects.toThrow(/StockNumbers/);
    await writeFile(configPath, '{"StockNumbers":"2330"}', 'utf8');
    const review = await inspectLegacyConfig(configPath);
    await expect(syncLegacyStockNumbers({ configPath, reviewedHash: review.hash, activeStockCodes: ['oops'] }))
      .rejects.toThrow(/股票代號/);
    await expect(syncLegacyStockNumbers({ configPath, reviewedHash: review.hash, activeStockCodes: [] }))
      .rejects.toThrow(/至少一間/);
  });
});
