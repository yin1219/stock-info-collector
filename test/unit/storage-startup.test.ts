import { describe, expect, it, vi } from 'vitest';
import { openStorageBeforeServices } from '../../src/main/storage-startup';

describe('local-data-management / startup migration gate', () => {
  it('does not start background services when opening or migrating storage fails', () => {
    const startServices = vi.fn();
    const migrationError = new Error('injected migration failure');

    expect(() => openStorageBeforeServices(() => { throw migrationError; }, startServices))
      .toThrow(migrationError);
    expect(startServices).not.toHaveBeenCalled();
  });

  it('starts services only after migrations have opened storage successfully', () => {
    const database = { close: vi.fn() };
    const order: string[] = [];
    const result = openStorageBeforeServices(() => { order.push('migrated'); return database; }, () => order.push('services'));

    expect(order).toEqual(['migrated', 'services']);
    expect(result).toBe(database);
  });
});
