const path = require('node:path');
const { build } = require('vite');
const createJiti = require('jiti');
const ViteConfigGenerator = require('@electron-forge/plugin-vite/dist/ViteConfig').default;

async function main() {
  const root = path.resolve(__dirname, '..');
  const jiti = createJiti(__filename);
  const forge = (await jiti.import(path.join(root, 'forge.config.ts'))).default;
  const plugin = forge.plugins.find((candidate) => candidate.constructor.name === 'VitePlugin');
  if (!plugin) throw new Error('Electron Forge Vite plugin is missing from forge.config.ts');

  const generator = new ViteConfigGenerator(plugin.config, root, true);
  const configs = [
    ...await generator.getBuildConfigs(),
    ...await generator.getRendererConfig(),
  ];
  for (const config of configs) await build({ configFile: false, logLevel: 'error', ...config });
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`);
  process.exitCode = 1;
});
