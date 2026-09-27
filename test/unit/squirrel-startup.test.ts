import { describe, expect, it } from 'vitest';
import { getSquirrelShortcutCommand } from '../../src/main/squirrel-startup';

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
});
