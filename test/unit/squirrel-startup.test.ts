import { describe, expect, it } from 'vitest';
import { cleanupConferenceTaskOnUninstall, getSquirrelShortcutCommand } from '../../src/main/squirrel-startup';

describe('windows-distribution / shortcuts', () => {
  it('creates desktop and Start menu shortcuts on install and update', () => {
    for (const event of ['--squirrel-install', '--squirrel-updated']) {
      expect(getSquirrelShortcutCommand(event, 'StockReporterAssistant.exe', 'C:/App/Update.exe'))
        .toEqual({ executable: 'C:/App/Update.exe', args: ['--createShortcut=StockReporterAssistant.exe', '--shortcut-locations=Desktop,StartMenu'] });
    }
  });

  it('removes shortcuts on uninstall and ignores obsolete versions', () => {
    expect(getSquirrelShortcutCommand('--squirrel-uninstall', 'StockReporterAssistant.exe', 'C:/App/Update.exe'))
      .toEqual({ executable: 'C:/App/Update.exe', args: ['--removeShortcut=StockReporterAssistant.exe'] });
    expect(getSquirrelShortcutCommand('--squirrel-obsolete', 'StockReporterAssistant.exe', 'C:/App/Update.exe'))
      .toBeNull();
  });

  it('asks to remove only a verified v2 task on uninstall, never on update', async () => {
    const removed: string[] = [];
    const gateway = { async removeIfOwned(expectedUpdateExe: string) {
      removed.push(expectedUpdateExe);
      return 'removed' as const;
    } };
    expect(await cleanupConferenceTaskOnUninstall('--squirrel-updated', 'C:\\App\\Update.exe', gateway)).toBe('untouched');
    expect(removed).toEqual([]);
    expect(await cleanupConferenceTaskOnUninstall('--squirrel-uninstall', 'C:\\App\\Update.exe', gateway)).toBe('removed');
    expect(removed).toEqual(['C:\\App\\Update.exe']);
  });
});
