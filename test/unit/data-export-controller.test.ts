import { describe, expect, it, vi } from 'vitest';
import { createDataExportController } from '../../src/main/data-export-controller';

describe('local-data-management / export controller', () => {
  it('exports to the user-selected destination and reports the completed path', async () => {
    const controller = createDataExportController({
      selectDestination: vi.fn(async () => 'C:/Users/test/Documents/reporter.json'),
      exportTo: vi.fn(async (destination) => destination),
    });

    await expect(controller.exportUserData()).resolves.toEqual({
      status: 'exported',
      destination: 'C:/Users/test/Documents/reporter.json',
    });
  });

  it('does not export when the save dialog is cancelled', async () => {
    const exportTo = vi.fn();
    const controller = createDataExportController({ selectDestination: vi.fn(async () => null), exportTo });

    await expect(controller.exportUserData()).resolves.toEqual({ status: 'cancelled' });
    expect(exportTo).not.toHaveBeenCalled();
  });
});
