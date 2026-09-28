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

test('windows-distribution / installer / keeps the on-disk executable ASCII while preserving the Chinese product name', async () => {
  const createJiti = require('jiti');
  const forge = (await createJiti(__filename).import(path.join(root, 'forge.config.ts'))).default;
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const packagedSmoke = await readFile(path.join(root, 'scripts/packaged-windows-smoke.cjs'), 'utf8');

  assert.equal(packageJson.productName, '股市記者小幫手');
  assert.equal(forge.packagerConfig.executableName, 'StockReporterAssistant');
  assert.match(packagedSmoke, /StockReporterAssistant\.exe/);
});

test('windows-distribution / installer / exposes a repeatable Squirrel archive smoke check', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  assert.match(packageJson.scripts['test:installer-smoke'], /installer-package-smoke\.ps1/);
  await access(path.join(root, 'scripts', 'installer-package-smoke.ps1'));
});

test('windows-distribution / upgrade / advances beyond the installed 1.0.0 release with matching lockfile metadata', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const lock = JSON.parse(await readFile(path.join(root, 'package-lock.json'), 'utf8'));
  assert.notEqual(packageJson.version, '1.0.0');
  assert.equal(lock.version, packageJson.version);
  assert.equal(lock.packages[''].version, packageJson.version);
});

test('windows-distribution / upgrade / uses a newer installer version than the installed 1.0.2 with the garbled executable name', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const lock = JSON.parse(await readFile(path.join(root, 'package-lock.json'), 'utf8'));
  const [major, minor, patch] = packageJson.version.split('.').map(Number);

  assert.ok(major > 1 || (major === 1 && (minor > 0 || patch > 2)));
  assert.equal(lock.version, packageJson.version);
  assert.equal(lock.packages[''].version, packageJson.version);
});

test('windows-distribution / upgrade / increments beyond the installed 1.0.3 before packaging review fixes', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const [major, minor, patch] = packageJson.version.split('.').map(Number);
  assert.ok(major > 1 || (major === 1 && (minor > 0 || patch > 3)));
});

test('windows-distribution / upgrade / increments beyond installed 1.0.4 for the scheduler correction', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const lock = JSON.parse(await readFile(path.join(root, 'package-lock.json'), 'utf8'));
  const [major, minor, patch] = packageJson.version.split('.').map(Number);
  assert.ok(major > 1 || (major === 1 && (minor > 0 || patch > 4)));
  assert.equal(lock.version, packageJson.version);
  assert.equal(lock.packages[''].version, packageJson.version);
});

test('windows-distribution / upgrade / increments beyond installed 1.0.5 for calendar disconnect and settings UI fixes', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const lock = JSON.parse(await readFile(path.join(root, 'package-lock.json'), 'utf8'));
  const [major, minor, patch] = packageJson.version.split('.').map(Number);
  assert.ok(major > 1 || (major === 1 && (minor > 0 || patch > 5)));
  assert.equal(lock.version, packageJson.version);
  assert.equal(lock.packages[''].version, packageJson.version);
});

test('material-event-monitoring / live source / uses dated MOPS website search as a degraded source', async () => {
  const main = await readFile(path.join(root, 'src/main/main.ts'), 'utf8');
  assert.match(main, /primary:\s*createMopsSearchMaterialProvider\(http\)/);
  assert.match(main, /async post\(url: string, body: string\)/);
});

test('windows-distribution / installer / embeds the application icon in Forge output', async () => {
  const forge = await readFile(path.join(root, 'forge.config.ts'), 'utf8');
  await access(path.join(root, 'assets', 'app-icon.ico'));
  assert.match(forge, /icon:\s*['"]assets\/app-icon['"]/);
  assert.match(forge, /setupIcon:\s*['"]assets\/app-icon\.ico['"]/);
});

test('desktop-app-lifecycle / isolated test mode / does not keep an operating-system tray alive', async () => {
  const main = await readFile(path.join(root, 'src/main/main.ts'), 'utf8');
  assert.match(main, /process\.env\.REPORTER_USER_DATA_DIR\s*&&\s*process\.env\.REPORTER_TEST_TRAY\s*===\s*'1'/);
  assert.match(main, /if \(!process\.env\.REPORTER_USER_DATA_DIR\s*\|\|\s*enableIsolatedTrayHarness\)[\s\S]*new ElectronTray/);
});

test('test infrastructure / Electron E2E / rebuilds the isolated Forge Vite shell before launch', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  await access(path.join(root, 'scripts', 'build-e2e-shell.cjs'));
  assert.match(packageJson.scripts['test:e2e'], /build-e2e-shell\.cjs.*playwright test/);
});

test('desktop-app-lifecycle / native modules / leaves better-sqlite3 loadable from packaged node_modules', async () => {
  const config = await readFile(path.join(root, 'vite.main.config.mts'), 'utf8');
  const forge = await readFile(path.join(root, 'forge.config.ts'), 'utf8');
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  assert.match(config, /external:\s*\[[^\]]*better-sqlite3/s);
  assert.match(forge, /node_modules/);
  assert.match(forge, /AutoUnpackNativesPlugin/);
  assert.equal(packageJson.devDependencies['@electron-forge/plugin-auto-unpack-natives'], '7.11.2');
});

test('windows-distribution / installer / packages only the native dependency left external by Vite', async () => {
  const createJiti = require('jiti');
  const forge = (await createJiti(__filename).import(path.join(root, 'forge.config.ts'))).default;
  const ignored = forge.packagerConfig.ignore;
  assert.equal(ignored(path.join(root, 'node_modules', 'better-sqlite3', 'lib', 'index.js')), false);
  assert.equal(ignored(path.join(root, 'node_modules', 'googleapis', 'build', 'src', 'index.js')), true);
  assert.equal(ignored(path.join(root, 'node_modules', 'axios', 'index.js')), true);
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
  const forbiddenImport = /(?:\bfrom\s*|\bimport\s*\(?\s*|\brequire\s*\(\s*)['"](?:node:|electron['"]|better-sqlite3|sqlite3|(?:\.\.\/)+(?:src\/)?(?:main|providers|repositories|services)(?:\/|['"]))/;
  for (const specifier of ['../main/main', '../services/scheduler', '../../providers/mops-search-material', '../repositories/index']) {
    assert.match(`import x from '${specifier}'`, forbiddenImport, `must reject ${specifier}`);
  }
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
