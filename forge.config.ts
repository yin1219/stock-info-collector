import type { ForgeConfig } from '@electron-forge/shared-types';
import { VitePlugin } from '@electron-forge/plugin-vite';
import { MakerSquirrel } from '@electron-forge/maker-squirrel';
import { AutoUnpackNativesPlugin } from '@electron-forge/plugin-auto-unpack-natives';
import { existsSync } from 'node:fs';
import path from 'node:path';

const oauthClientResource = process.env.REPORTER_GOOGLE_OAUTH_CLIENT_FILE
  ? path.resolve(process.env.REPORTER_GOOGLE_OAUTH_CLIENT_FILE)
  : undefined;
if (oauthClientResource && (path.basename(oauthClientResource) !== 'google-oauth-client.json' || !existsSync(oauthClientResource))) {
  throw new Error('發佈用 Google OAuth client 必須是已存在的 google-oauth-client.json');
}

const config: ForgeConfig = {
  outDir: 'artifacts/forge-out',
  packagerConfig: {
    asar: true,
    icon: 'assets/app-icon',
    executableName: 'StockReporterAssistant',
    electronZipDir: process.env.REPORTER_ELECTRON_ZIP_DIR || undefined,
    ...(oauthClientResource ? { extraResource: [oauthClientResource] } : {}),
    ignore: (file: string) => {
      if (!file) return false;
      const normalized = file.replace(/\\/g, '/').replace(/^\/+/, '');
      const root = `${process.cwd().replace(/\\/g, '/').replace(/\/$/, '')}/`;
      const relative = normalized.startsWith(root) ? normalized.slice(root.length) : normalized;
      if (!relative) return false;
      return !(relative === 'package.json' || relative === '.vite' || relative.startsWith('.vite/')
        || relative === 'assets' || relative.startsWith('assets/')
        || relative === 'node_modules' || relative === 'node_modules/better-sqlite3'
        || relative.startsWith('node_modules/better-sqlite3/'));
    },
  },
  rebuildConfig: {},
  makers: [new MakerSquirrel({
    name: 'StockReporterAssistant',
    authors: 'Stock Reporter Assistant',
    description: '財經記者本機市場資訊監控與行事曆小幫手',
    setupExe: 'StockReporterAssistantSetup.exe',
    setupIcon: 'assets/app-icon.ico',
    noMsi: true,
  }, ['win32'])],
  plugins: [
    new AutoUnpackNativesPlugin({}),
    new VitePlugin({
      build: [
        { entry: 'src/main/main.ts', config: 'vite.main.config.mts', target: 'main' },
        { entry: 'src/preload/preload.ts', config: 'vite.preload.config.mts', target: 'preload' },
      ],
      renderer: [
        { name: 'main_window', config: 'vite.renderer.config.mts' },
      ],
    }),
  ],
};

export default config;
