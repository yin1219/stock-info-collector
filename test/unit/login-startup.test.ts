import { describe, expect, it, vi } from 'vitest';
import { createLoginStartupController } from '../../src/main/login-startup';

describe('desktop-app-lifecycle / login startup', () => {
  it('reads visible and hidden login entries and applies the packaged Squirrel launcher', () => {
    const get = vi.fn(({ args }: { args: string[] }) => ({ openAtLogin: args[0] === '--hidden' }));
    const set = vi.fn();
    const controller = createLoginStartupController({
      getLoginItemSettings: get, setLoginItemSettings: set,
    }, 'C:\\App\\app-1.0.0\\StockReporterAssistant.exe', true);

    expect(controller.getSettings()).toEqual({ enabled: true, startHidden: true });
    expect(get).toHaveBeenNthCalledWith(1, { path: 'C:\\App\\StockReporterAssistant.exe', args: [] });
    expect(get).toHaveBeenNthCalledWith(2, { path: 'C:\\App\\StockReporterAssistant.exe', args: ['--hidden'] });

    controller.saveSettings({ enabled: true, startHidden: true });
    expect(set).toHaveBeenCalledWith({
      openAtLogin: true, enabled: true, path: 'C:\\App\\StockReporterAssistant.exe', args: ['--hidden'],
    });
  });

  it('removes both possible startup argument entries when login launch is disabled', () => {
    const set = vi.fn();
    const controller = createLoginStartupController({ getLoginItemSettings: () => ({ openAtLogin: false }), setLoginItemSettings: set }, 'C:\\App\\app-1.0.0\\StockReporterAssistant.exe', true);
    controller.saveSettings({ enabled: false, startHidden: false });
    expect(set).toHaveBeenNthCalledWith(1, { openAtLogin: false, enabled: true, path: 'C:\\App\\StockReporterAssistant.exe', args: [] });
    expect(set).toHaveBeenNthCalledWith(2, { openAtLogin: false, enabled: true, path: 'C:\\App\\StockReporterAssistant.exe', args: ['--hidden'] });
  });

  it('does not change OS login settings from an unpackaged development run', () => {
    const set = vi.fn();
    const controller = createLoginStartupController({ getLoginItemSettings: () => ({ openAtLogin: false }), setLoginItemSettings: set }, 'C:\\repo\\electron.exe', false);
    expect(() => controller.saveSettings({ enabled: true, startHidden: true })).toThrow('必須先安裝 Windows 版應用程式');
    expect(set).not.toHaveBeenCalled();
  });

  it('does not claim a successful save when Windows does not register the login item', () => {
    const controller = createLoginStartupController({
      getLoginItemSettings: () => ({ openAtLogin: false }), setLoginItemSettings: vi.fn(),
    }, 'C:\\App\\app-1.0.3\\StockReporterAssistant.exe', true);
    expect(() => controller.saveSettings({ enabled: true, startHidden: false })).toThrow('登入啟動設定未生效');
  });

  it('reports a Windows-disabled startup item separately from an unchecked preference', () => {
    const controller = createLoginStartupController({
      getLoginItemSettings: () => ({ openAtLogin: true, executableWillLaunchAtLogin: false }),
      setLoginItemSettings: vi.fn(),
    }, 'C:\\App\\app-1.0.3\\StockReporterAssistant.exe', true);
    expect(controller.getSettings()).toEqual({ enabled: true, startHidden: true, blockedByWindows: true });
  });
});
