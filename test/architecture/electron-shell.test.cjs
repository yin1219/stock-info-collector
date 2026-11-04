const { access, readdir, readFile } = require('node:fs/promises');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
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

test('desktop-app-lifecycle / login startup defaults / only initializes a new formal profile', async () => {
  const main = await readFile(path.join(root, 'src/main/main.ts'), 'utf8');
  assert.match(main, /const isNewProfile = !existsSync\(databasePath\)/);
  assert.match(main, /isNewProfile && !process\.env\.REPORTER_USER_DATA_DIR/);
  assert.match(main, /loginStartup\.applyNewProfileDefaults\(\)/);
});

test('desktop-app-lifecycle / development preview launcher / offers an Explorer entry with isolated data', async () => {
  const cmd = await readFile(path.join(root, 'scripts/launch-dev-preview.cmd'), 'utf8');
  const script = await readFile(path.join(root, 'scripts/launch-dev-preview.ps1'), 'utf8');
  assert.ok([...script].every((character) => character.charCodeAt(0) < 128), 'Windows PowerShell 5.1 must parse the preview launcher without a UTF-8 BOM');
  assert.match(cmd, /launch-dev-preview\.ps1/);
  assert.match(script, /\$PSScriptRoot/);
  assert.match(script, /REPORTER_USER_DATA_DIR/);
  assert.match(script, /REPORTER_TEST_TRAY/);
  assert.match(script, /npm(?:\.cmd)?/i);
  assert.doesNotMatch(cmd, /EnableGoogleOAuth/);
  assert.match(script, /if \(\$EnableGoogleOAuth\) \{ \$env:REPORTER_TEST_LIVE_OAUTH = '1' \}/);
});

test('desktop-app-lifecycle / development preview launcher / parses under Windows PowerShell 5.1', () => {
  if (process.platform !== 'win32') return;
  const result = spawnSync('powershell.exe', [
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
    path.join(root, 'scripts/launch-dev-preview.ps1'), '-ValidateOnly',
  ], { encoding: 'utf8' });
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /DEV_PREVIEW_LAUNCHER_OK/);
});

test('windows-distribution / conference task gateway / uses create-only Task Scheduler registration', async () => {
  const scriptPath = path.join(root, 'scripts', 'conference-task.ps1');
  const script = await readFile(scriptPath, 'utf8');
  assert.ok([...script].every((character) => character.charCodeAt(0) < 128), 'PowerShell 5.1 gateway must remain ASCII without a BOM');
  assert.match(script, /RegisterTaskDefinition\(\$TaskName,\s*\$definition,\s*2,/);
  assert.match(script, /MultipleInstances\s*=\s*2/);
  assert.match(script, /LogonType\s*=\s*3/);
  assert.doesNotMatch(script, /UnregisterTask|TASK_CREATE_OR_UPDATE/i);
  assert.match(script, /if \(\$Action -eq 'Remove'\)[\s\S]*\$folder\.DeleteTask\(\$TaskName, 0\)/);
  assert.match(script, /\$taskAction\.Arguments -ne '\-\-processStart StockReporterAssistant\.exe \-\-process-start-args "\-\-conference-scheduled"'/);
  if (process.platform !== 'win32') return;
  const result = spawnSync('powershell.exe', [
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath, '-ValidateOnly',
  ], { encoding: 'utf8' });
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /CONFERENCE_TASK_GATEWAY_OK/);
  assert.match(result.stdout, /CONFERENCE_TASK_MISSING_CODES_OK/);
  assert.match(result.stdout, /CONFERENCE_TASK_SINGLE_WRITER_OK/);
  assert.match(result.stdout, /CONFERENCE_TASK_DISABLED_LEGACY_OK/);
});

test('windows-distribution / conference task gateway / bundles the Task Scheduler script as a resource', async () => {
  const createJiti = require('jiti');
  const forge = (await createJiti(__filename).import(path.join(root, 'forge.config.ts'))).default;
  assert.ok(forge.packagerConfig.extraResource?.some((item) =>
    path.basename(item) === 'conference-task.ps1'));
});

test('conference-calendar-sync / packaged smoke / exercises the headless daily entry with isolated data', async () => {
  const smoke = await readFile(path.join(root, 'scripts/packaged-windows-smoke.cjs'), 'utf8');
  assert.match(smoke, /--conference-scheduled/);
  assert.match(smoke, /conference-scheduled-not-ready/);
  assert.match(smoke, /renderer-load-started/);
});

test('windows-distribution / uninstall / handles verified v2 task cleanup before shortcut removal without needing the app lock', async () => {
  const main = await readFile(path.join(root, 'src/main/main.ts'), 'utf8');
  assert.match(main, /cleanupConferenceTaskOnUninstall\(squirrelEvent/);
  assert.match(main, /isSquirrelStartup\s*\?\s*false\s*:\s*app\.requestSingleInstanceLock\(\)/);
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

test('windows-distribution / upgrade / advances beyond 1.0.6 for corrected reconciliation announcement timestamps', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const lock = JSON.parse(await readFile(path.join(root, 'package-lock.json'), 'utf8'));
  const [major, minor, patch] = packageJson.version.split('.').map(Number);
  assert.ok(major > 1 || (major === 1 && (minor > 0 || patch > 6)));
  assert.equal(lock.version, packageJson.version);
  assert.equal(lock.packages[''].version, packageJson.version);
});

test('windows-distribution / upgrade / advances beyond installed 1.0.7 for the official conference redirect host', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const lock = JSON.parse(await readFile(path.join(root, 'package-lock.json'), 'utf8'));
  const [major, minor, patch] = packageJson.version.split('.').map(Number);
  assert.ok(major > 1 || (major === 1 && (minor > 0 || patch > 7)));
  assert.equal(lock.version, packageJson.version);
  assert.equal(lock.packages[''].version, packageJson.version);
});

test('windows-distribution / upgrade / advances beyond installed 1.0.8 for the bundled daily conference job', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const lock = JSON.parse(await readFile(path.join(root, 'package-lock.json'), 'utf8'));
  const [major, minor, patch] = packageJson.version.split('.').map(Number);
  assert.ok(major > 1 || (major === 1 && (minor > 0 || patch > 8)));
  assert.equal(lock.version, packageJson.version);
  assert.equal(lock.packages[''].version, packageJson.version);
});

test('windows-distribution / upgrade / advances beyond installed 1.0.10 for disabled-legacy Calendar acceptance', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const lock = JSON.parse(await readFile(path.join(root, 'package-lock.json'), 'utf8'));
  const [major, minor, patch] = packageJson.version.split('.').map(Number);
  assert.ok(major > 1 || (major === 1 && (minor > 0 || patch > 10)));
  assert.equal(lock.version, packageJson.version);
  assert.equal(lock.packages[''].version, packageJson.version);
});

test('windows-distribution / upgrade / advances beyond installed 1.0.11 for login startup defaults', async () => {
  const packageJson = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
  const lock = JSON.parse(await readFile(path.join(root, 'package-lock.json'), 'utf8'));
  const [major, minor, patch] = packageJson.version.split('.').map(Number);
  assert.ok(major > 1 || (major === 1 && (minor > 0 || patch > 11)));
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
