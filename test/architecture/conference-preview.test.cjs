const { test } = require('node:test');
const assert = require('node:assert/strict');
const { readFile } = require('node:fs/promises');
const path = require('node:path');

const root = path.join(__dirname, '..', '..');

test('opt-in conference preview prints results and has no Google Calendar dependency', async () => {
  const [config, canary] = await Promise.all([
    readFile(path.join(root, 'vitest.manual.config.mts'), 'utf8'),
    readFile(path.join(root, 'test', 'manual', 'live-conference-source.test.ts'), 'utf8'),
  ]);
  assert.match(config, /disableConsoleIntercept:\s*true/);
  assert.match(canary, /previewConferenceEvents/);
  assert.doesNotMatch(canary, /googleapis|calendar\.events|\.insert\(/i);
});
