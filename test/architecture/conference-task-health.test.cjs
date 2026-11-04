const { test } = require('node:test');
const assert = require('node:assert/strict');
const { readFile } = require('node:fs/promises');
const { execFileSync } = require('node:child_process');
const path = require('node:path');

test('read-only Task Scheduler inspection reports the previous result without running a task', async () => {
  const script = await readFile(path.join(__dirname, '..', '..', 'scripts', 'conference-task.ps1'), 'utf8');
  const inspect = script.split("if ($Action -eq 'Inspect') {")[1]?.split("if ($TaskName -ne 'StockReporterAssistant-ConferenceDaily')")[0];
  assert.ok(inspect, 'Inspect action must exist');
  assert.match(inspect, /lastTaskResult\s*=/);
  assert.doesNotMatch(inspect, /RegisterTaskDefinition|DeleteTask|\.Run\(/);
});

test('registration rechecks for a v1 task immediately before creating the v2 writer', async () => {
  const script = await readFile(path.join(__dirname, '..', '..', 'scripts', 'conference-task.ps1'), 'utf8');
  const register = script.split("if ($Action -eq 'Remove') {")[1]?.split('$folder.RegisterTaskDefinition(')[0];
  assert.ok(register, 'Register action must exist');
  assert.match(register, /Assert-LegacyTaskAbsent \$folder/);
  const guard = script.split('function Assert-LegacyTaskAbsent($Folder) {')[1]?.split('if ($ValidateOnly) {')[0];
  assert.ok(guard, 'single-writer guard must exist');
  assert.match(guard, /foreach \(\$legacyName in @\('stock-info-collector-daily-work', 'stock-info-collector'\)\)/);
  assert.match(guard, /\$Folder\.GetTask\(\$legacyName\)/);
  assert.match(guard, /Test-MissingTaskException/);
  assert.match(script, /\[switch\]\$AllowDisabledLegacy/);
  assert.match(guard, /\$AllowDisabledLegacy/);
  assert.match(guard, /\$legacyTask\.State\s*-ne\s*1/);
});

test('v2 removal recognizes the current user even when Task Scheduler returns a short account name', async () => {
  const script = await readFile(path.join(__dirname, '..', '..', 'scripts', 'conference-task.ps1'), 'utf8');
  const remove = script.split("if ($Action -eq 'Remove') {")[1]?.split("if (-not [IO.Path]::IsPathRooted($Executable) -or\n    [IO.Path]::GetFileName")[0];
  assert.ok(remove, 'Remove action must exist');
  assert.match(remove, /Test-CurrentUserPrincipal\s+\(\[string\]\$definition\.Principal\.UserId\)/);
  const output = execFileSync('powershell.exe', [
    '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
    '-File', path.join(__dirname, '..', '..', 'scripts', 'conference-task.ps1'),
    '-ValidateOnly',
  ], { encoding: 'utf8', timeout: 15_000 });
  assert.match(output, /CONFERENCE_TASK_OWNER_IDENTITY_OK/);
});

test('owned v2 17:45 jobs migrate with a backup; new jobs use 19:00 and removal accepts both owned times', async () => {
  const script = await readFile(path.join(__dirname, '..', '..', 'scripts', 'conference-task.ps1'), 'utf8');
  assert.match(script, /ValidateSet\('Inspect', 'Register', 'Remove', 'MigrateTime'\)/);
  assert.match(script, /\$trigger\.StartBoundary\s*=.*T19:00:00/);
  const migrate = script.split("if ($Action -eq 'MigrateTime') {")[1]?.split("if ($Action -eq 'Remove') {")[0];
  assert.ok(migrate, 'MigrateTime action must exist');
  assert.match(migrate, /Assert-LegacyTaskAbsent \$folder/);
  assert.match(migrate, /Test-CurrentUserPrincipal/);
  assert.match(migrate, /\[IO\.File\]::WriteAllText/);
  assert.match(migrate, /T17:45:00/);
  assert.match(migrate, /T19:00:00/);
  assert.match(migrate, /RegisterTaskDefinition/);
  assert.match(migrate, /RegisterTask\(/);
  assert.match(migrate, /\[\+\-\]/, 'time migration must retain an explicit timezone offset');
  const remove = script.split("if ($Action -eq 'Remove') {")[1]?.split("if (-not [IO.Path]::IsPathRooted($Executable) -or\n    [IO.Path]::GetFileName")[0];
  assert.match(remove, /T17:45:00/);
  assert.match(remove, /T19:00:00/);
});
