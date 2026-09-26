import type { ForgeConfig } from '@electron-forge/shared-types';
import { VitePlugin } from '@electron-forge/plugin-vite';
import { MakerSquirrel } from '@electron-forge/maker-squirrel';

const config: ForgeConfig = {
  outDir: 'artifacts/forge-out',
  packagerConfig: {
    asar: true,
  },
  rebuildConfig: {},
  makers: [new MakerSquirrel({
    name: 'StockReporterAssistant',
    authors: 'Stock Reporter Assistant',
    description: '財經記者本機市場資訊監控與行事曆小幫手',
    setupExe: 'StockReporterAssistantSetup.exe',
    noMsi: true,
  }, ['win32'])],
  plugins: [
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
