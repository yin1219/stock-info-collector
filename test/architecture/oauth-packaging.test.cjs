const { mkdtemp, rm, writeFile, readFile } = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');

const root = path.resolve(__dirname, '..', '..');

test('conference-calendar-sync / packaged OAuth client / copies only an explicitly supplied application client resource', async () => {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'stock-reporter-oauth-build-test-'));
  const clientPath = path.join(directory, 'google-oauth-client.json');
  await writeFile(clientPath, '{}', 'utf8');
  const previous = process.env.REPORTER_GOOGLE_OAUTH_CLIENT_FILE;
  try {
    process.env.REPORTER_GOOGLE_OAUTH_CLIENT_FILE = clientPath;
    const createJiti = require('jiti');
    const forge = (await createJiti(__filename).import(path.join(root, 'forge.config.ts'))).default;
    assert.deepEqual(forge.packagerConfig.extraResource, [clientPath]);
    const main = await readFile(path.join(root, 'src', 'main', 'main.ts'), 'utf8');
    assert.match(main, /path\.join\(process\.resourcesPath, 'google-oauth-client\.json'\)/);
  } finally {
    if (previous === undefined) delete process.env.REPORTER_GOOGLE_OAUTH_CLIENT_FILE;
    else process.env.REPORTER_GOOGLE_OAUTH_CLIENT_FILE = previous;
    await rm(directory, { recursive: true, force: true });
  }
});

test('conference-calendar-sync / packaged OAuth client / keeps maintainer client files out of Git by default', async () => {
  const ignore = await readFile(path.join(root, '.gitignore'), 'utf8');
  assert.match(ignore, /^google-oauth-client\.json$/m);
});

test('conference-calendar-sync / packaged OAuth acceptance / targets the packaged executable in an isolated profile', async () => {
  const script = await readFile(path.join(root, 'scripts', 'manual-google-oauth-acceptance.cjs'), 'utf8');
  assert.match(script, /process\.env\.REPORTER_PACKAGED_EXECUTABLE/);
  assert.match(script, /executablePath: path\.resolve\(process\.env\.REPORTER_PACKAGED_EXECUTABLE\)/);
  assert.match(script, /window\.bringToFront\(\)/);
  assert.match(script, /REPORTER_USER_DATA_DIR: userData/);
});
