const { access, readdir, readFile } = require('node:fs/promises');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');

const root = path.resolve(__dirname, '..', '..');

test('desktop-app-lifecycle / shell / pins the Electron Forge Vite React TypeScript toolchain', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));

  assert.equal(packageJson.main, '.vite/build/main.js');
  assert.equal(packageJson.engines.node, '22.x');
  for (const dependency of [
    'electron',
    '@electron-forge/cli',
    '@electron-forge/plugin-vite',
    'vite',
    'typescript',
    'react',
    'react-dom',
  ]) {
    const version = packageJson.dependencies?.[dependency]
      ?? packageJson.devDependencies?.[dependency];
    assert.match(version, /^\d+\.\d+\.\d+$/, `${dependency} must use an exact version`);
  }
});

test('desktop-app-lifecycle / shell / contains isolated main, preload and renderer entries', async () => {
  await Promise.all([
    'forge.config.ts',
    'vite.main.config.mts',
    'vite.preload.config.mts',
    'vite.renderer.config.mts',
    'src/main/main.ts',
    'src/preload/preload.ts',
    'src/renderer/index.html',
    'src/renderer/main.tsx',
  ].map((relativePath) => access(path.join(root, relativePath))));
});

test('windows-distribution / installer / configures Squirrel per-user setup and handles Squirrel lifecycle events', async () => {
  const forge = await readFile(path.join(root, 'forge.config.ts'), 'utf8');
  const main = await readFile(path.join(root, 'src/main/main.ts'), 'utf8');
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  assert.match(forge, /MakerSquirrel/);
  assert.match(forge, /setupExe/);
  assert.match(forge, /noMsi:\s*true/);
  assert.match(main, /--squirrel-(?:install|updated|uninstall|obsolete)/);
  assert.match(packageJson.devDependencies['@electron-forge/maker-squirrel'], /^7\.11\.2$/);
});

test('desktop-app-lifecycle / native modules / leaves better-sqlite3 loadable from packaged node_modules', async () => {
  const config = await readFile(path.join(root, 'vite.main.config.mts'), 'utf8');
  assert.match(config, /external:\s*\[[^\]]*better-sqlite3/s);
});

test('renderer architecture / has domain, provider, repository and service boundaries without privileged imports', async () => {
  for (const relativePath of [
    'src/domain',
    'src/providers',
    'src/repositories',
    'src/services',
  ]) {
    await access(path.join(root, relativePath));
  }

  async function collectSourceFiles(directory) {
    const entries = await readdir(directory, { withFileTypes: true });
    const nested = await Promise.all(entries.map(async (entry) => {
      const entryPath = path.join(directory, entry.name);
      if (entry.isDirectory()) return collectSourceFiles(entryPath);
      return /\.tsx?$/.test(entry.name) ? [entryPath] : [];
    }));
    return nested.flat();
  }

  const rendererFiles = await collectSourceFiles(path.join(root, 'src/renderer'));
  const forbiddenImport = /(?:\bfrom\s*|\bimport\s*\(?\s*|\brequire\s*\(\s*)['"](?:node:|electron['"]|better-sqlite3|sqlite3|(?:\.\.\/)+src\/(?:main|providers|repositories|services)(?:\/|['"]))/;
  for (const filePath of rendererFiles) {
    const source = await readFile(filePath, 'utf8');
    assert.doesNotMatch(source, forbiddenImport, path.relative(root, filePath));
  }
});

test('security architecture / preload does not expose token or filesystem access to renderer', async () => {
  const preload = await readFile(path.join(root, 'src/preload/preload.ts'), 'utf8');
  assert.match(preload, /contextBridge\.exposeInMainWorld\('reporterApi'/);
  assert.doesNotMatch(preload, /safeStorage|oauth-token|accessToken|refreshToken|readFile|writeFile/);
});

test('security architecture / renderer declares a restrictive content security policy', async () => {
  const html = await readFile(path.join(root, 'src/renderer/index.html'), 'utf8');
  assert.match(html, /http-equiv="Content-Security-Policy"/);
  assert.match(html, /default-src 'self'/);
  assert.match(html, /script-src 'self'/);
  assert.doesNotMatch(html, /script-src[^>]*(?:unsafe-eval|https:\/\/)/);
});
