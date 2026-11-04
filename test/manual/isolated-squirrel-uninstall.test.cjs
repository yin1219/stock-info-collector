const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

test('windows-distribution / packaged Squirrel uninstall / removes only an isolated owned task', { skip: process.platform !== 'win32' }, () => {
  const script = path.resolve(__dirname, '../../scripts/isolated-squirrel-uninstall-smoke.ps1');
  const result = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script], {
    cwd: path.resolve(__dirname, '../..'), encoding: 'utf8', timeout: 120_000,
  });
  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
  assert.match(result.stdout, /PACKAGED_SQUIRREL_UNINSTALL_ISOLATED_OK/);
});
