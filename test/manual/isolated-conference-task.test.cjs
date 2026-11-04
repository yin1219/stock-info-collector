const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

test('windows-distribution / isolated Task Scheduler smoke / registers, inspects and removes only a temporary v2 task', { skip: process.platform !== 'win32' }, () => {
  const script = path.resolve(__dirname, '../../scripts/isolated-conference-task-smoke.ps1');
  const result = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script], {
    cwd: path.resolve(__dirname, '../..'), encoding: 'utf8', timeout: 90_000,
  });
  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
  assert.match(result.stdout, /CONFERENCE_TASK_WINDOWS_ISOLATED_OK/);
  assert.match(result.stdout, /CONFERENCE_TASK_FOREIGN_PRESERVED_OK/);
  assert.match(result.stdout, /CONFERENCE_TASK_ALTERNATE_V1_BLOCKED_OK/);
  assert.match(result.stdout, /CONFERENCE_TASK_ALTERNATE_V1_MIGRATION_BLOCKED_OK/);
  assert.match(result.stdout, /CONFERENCE_TASK_DISABLED_ALTERNATE_TRAILING_PATH_OK/);
});
