import { mkdir, mkdtemp, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { createSecretStore } from '../../src/services/secret-store';
import { createLogger } from '../../src/services/logger';

const temporaryRoots: string[] = [];

afterEach(async () => {
  await Promise.all(temporaryRoots.splice(0).map((directory) => rm(directory, { recursive: true, force: true })));
});

async function createTemporaryRoot(): Promise<string> {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-security-'));
  temporaryRoots.push(directory);
  return directory;
}

describe('local-data-management / operating-system protected tokens', () => {
  it('refuses to persist a token when OS encryption is unavailable', async () => {
    const tokenPath = path.join(await createTemporaryRoot(), 'oauth-token.dat');
    const storage = {
      isEncryptionAvailable: () => false,
      encryptString: (value: string) => Buffer.from(value),
      decryptString: (value: Buffer) => value.toString(),
    };

    await expect(createSecretStore(storage, tokenPath).save('must-not-be-written'))
      .rejects.toThrow('安全儲存目前無法使用');
    await expect(readFile(tokenPath)).rejects.toMatchObject({ code: 'ENOENT' });
  });

  it('stores only encrypted bytes and decrypts them through the injected OS storage port', async () => {
    const tokenPath = path.join(await createTemporaryRoot(), 'oauth-token.dat');
    const storage = {
      isEncryptionAvailable: () => true,
      encryptString: (value: string) => Buffer.from(`protected:${value}`),
      decryptString: (value: Buffer) => value.toString().replace(/^protected:/, ''),
    };
    const store = createSecretStore(storage, tokenPath);

    await store.save('refresh-token-value');

    const storedBytes = await readFile(tokenPath);
    expect(storedBytes.toString('utf8')).not.toContain('refresh-token-value');
    expect(await store.load()).toBe('refresh-token-value');
  });

  it('returns no token for a missing file and refuses reads when OS encryption is unavailable', async () => {
    const tokenPath = path.join(await createTemporaryRoot(), 'missing-token.dat');
    const unavailable = createSecretStore({
      isEncryptionAvailable: () => false,
      encryptString: (value) => Buffer.from(value),
      decryptString: (value) => value.toString(),
    }, tokenPath);
    await expect(unavailable.load()).rejects.toThrow('拒絕讀取憑證');

    const available = createSecretStore({
      isEncryptionAvailable: () => true,
      encryptString: (value) => Buffer.from(value),
      decryptString: (value) => value.toString(),
    }, tokenPath);
    await expect(available.load()).resolves.toBeUndefined();

    const directoryPath = path.join(await createTemporaryRoot(), 'not-a-token-file');
    await mkdir(directoryPath);
    const store = createSecretStore({
      isEncryptionAvailable: () => true,
      encryptString: (value) => Buffer.from(value),
      decryptString: (value) => value.toString(),
    }, directoryPath);
    await expect(store.load()).rejects.toMatchObject({ code: expect.not.stringMatching('ENOENT') });
  });
});

describe('local-data-management / structured diagnostic logs', () => {
  it('persists actionable error details while masking token values from messages and fields', async () => {
    const directory = await createTemporaryRoot();
    const logger = createLogger({ directory, now: () => new Date('2026-09-26T00:00:00.000Z') });

    logger.error('google-calendar-auth-failed', new Error('Authorization: Bearer live-secret-value'), {
      accessToken: 'another-live-secret',
      stockCode: '2330',
    });

    const log = await readFile(path.join(directory, 'application.log'), 'utf8');
    expect(log).toContain('google-calendar-auth-failed');
    expect(log).toContain('2330');
    expect(log).toContain('[REDACTED]');
    expect(log).not.toContain('live-secret-value');
    expect(log).not.toContain('another-live-secret');
    logger.info('settings-loaded');
    logger.error('plain-diagnostic', { secret: 'must-not-leak', details: ['refresh_token=second-secret', 'ordinary text'] });
    const updatedLog = await readFile(path.join(directory, 'application.log'), 'utf8');
    expect(updatedLog).toContain('settings-loaded');
    expect(updatedLog).toContain('plain-diagnostic');
    expect(updatedLog).not.toContain('second-secret');
    expect(updatedLog).not.toContain('must-not-leak');
  });
});
