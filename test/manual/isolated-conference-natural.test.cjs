const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

test('windows-distribution / near-time daily trigger / naturally launches the isolated packaged headless entry', { skip: process.platform !== 'win32' }, () => {
  const script = path.resolve(__dirname, '../../scripts/isolated-conference-natural-smoke.ps1');
  const result = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script], {
    cwd: path.resolve(__dirname, '../..'), encoding: 'utf8', timeout: 180_000,
  });
  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
  assert.match(result.stdout, /CONFERENCE_DAILY_NATURAL_ISOLATED_OK/);
  assert.match(result.stdout, /triggerType=Daily/);
  assert.match(result.stdout, /formal-tasks-unchanged/);
});
