import { readFile, writeFile } from 'node:fs/promises';

export interface OperatingSystemStorage {
  isEncryptionAvailable(): boolean;
  encryptString(value: string): Buffer;
  decryptString(value: Buffer): string;
}

export function createSecretStore(storage: OperatingSystemStorage, filename: string) {
  return {
    async save(value: string): Promise<void> {
      if (!storage.isEncryptionAvailable()) {
        throw new Error('安全儲存目前無法使用，拒絕以明文儲存憑證');
      }
      const encrypted = storage.encryptString(value);
      await writeFile(filename, encrypted.toString('base64'), { encoding: 'utf8', mode: 0o600 });
    },

    async load(): Promise<string | undefined> {
      if (!storage.isEncryptionAvailable()) {
        throw new Error('安全儲存目前無法使用，拒絕讀取憑證');
      }
      let encryptedBase64: string;
      try {
        encryptedBase64 = await readFile(filename, 'utf8');
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'ENOENT') return undefined;
        throw error;
      }
      return storage.decryptString(Buffer.from(encryptedBase64, 'base64'));
    },
  };
}
