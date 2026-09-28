import path from 'node:path';

interface LoginItemSettingsPort {
  getLoginItemSettings(options: { path: string; args: string[] }): { openAtLogin: boolean; executableWillLaunchAtLogin?: boolean };
  setLoginItemSettings(options: { openAtLogin: boolean; enabled: boolean; path: string; args: string[] }): void;
}

export interface LoginStartupSettings {
  enabled: boolean;
  startHidden: boolean;
  blockedByWindows?: boolean;
}

export function createLoginStartupController(app: LoginItemSettingsPort, executablePath: string, packaged: boolean) {
  const launcherPath = path.resolve(path.dirname(executablePath), '..', path.basename(executablePath));
  const visibleArgs: string[] = [];
  const hiddenArgs = ['--hidden'];
  const read = (args: string[]) => app.getLoginItemSettings({ path: launcherPath, args });
  return {
    getSettings(): LoginStartupSettings {
      if (!packaged) return { enabled: false, startHidden: false };
      const visible = read(visibleArgs);
      const hidden = read(hiddenArgs);
      const enabled = visible.openAtLogin || hidden.openAtLogin;
      const blockedByWindows = (visible.openAtLogin && visible.executableWillLaunchAtLogin === false)
        || (hidden.openAtLogin && hidden.executableWillLaunchAtLogin === false);
      return { enabled, startHidden: hidden.openAtLogin, ...(blockedByWindows ? { blockedByWindows: true } : {}) };
    },
    saveSettings(settings: LoginStartupSettings): LoginStartupSettings {
      if (!settings || typeof settings.enabled !== 'boolean' || typeof settings.startHidden !== 'boolean') {
        throw new Error('登入啟動設定格式無效');
      }
      if (!packaged && settings.enabled) throw new Error('必須先安裝 Windows 版應用程式才能設定登入啟動');
      if (!packaged) return { enabled: false, startHidden: false };

      if (settings.enabled) {
        const activeArgs = settings.startHidden ? hiddenArgs : visibleArgs;
        const staleArgs = settings.startHidden ? visibleArgs : hiddenArgs;
        app.setLoginItemSettings({ openAtLogin: false, enabled: true, path: launcherPath, args: staleArgs });
        app.setLoginItemSettings({ openAtLogin: true, enabled: true, path: launcherPath, args: activeArgs });
      } else {
        app.setLoginItemSettings({ openAtLogin: false, enabled: true, path: launcherPath, args: visibleArgs });
        app.setLoginItemSettings({ openAtLogin: false, enabled: true, path: launcherPath, args: hiddenArgs });
      }
      const actual = this.getSettings();
      if (actual.enabled !== settings.enabled || (settings.enabled && actual.startHidden !== settings.startHidden)) {
        throw new Error('登入啟動設定未生效，請檢查 Windows 啟動應用程式設定');
      }
      return actual;
    },
  };
}
