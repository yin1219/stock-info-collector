const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const path = require('node:path');
const { test } = require('node:test');

test('desktop-app-lifecycle / Electron E2E / only isolated tray lifecycle cases request a native tray', () => {
  const source = readFileSync(path.join(__dirname, '..', 'e2e', 'electron.smoke.spec.ts'), 'utf8');
  const cases = source.split(/^test\(/m).slice(1);
  assert.ok(cases.length > 0);
  for (const item of cases) {
    if (item.startsWith("'desktop-app-lifecycle / isolated tray harness")
      || item.startsWith("'conference-calendar-sync / scheduled entry / reaches the existing background instance'")) continue;
    assert.doesNotMatch(item, /REPORTER_TEST_TRAY:\s*'1'/);
  }
});
