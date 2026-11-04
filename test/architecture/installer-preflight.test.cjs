const { test } = require('node:test');
const assert = require('node:assert/strict');
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');

const core = path.resolve(__dirname, '../../scripts/installer-preflight-core.ps1').replaceAll("'", "''");

test('windows-distribution / upgrade preflight does not leave a visible PowerShell console waiting for the app', () => {
  const preflight = fs.readFileSync(path.resolve(__dirname, '../../scripts/installer-preflight.ps1'), 'utf8');
  const wrapper = fs.readFileSync(path.resolve(__dirname, '../../scripts/wrap-windows-installer.cjs'), 'utf8');
  assert.match(wrapper, /AppLaunched=powershell\.exe -NoProfile -WindowStyle Hidden /);
  assert.match(wrapper, /\[Strings\]/);
  assert.match(wrapper, /FriendlyName=%FriendlyName%/);
  assert.match(wrapper, /FriendlyName=Stock Reporter Assistant Setup/);
  assert.match(preflight, /\$installer\.WaitForExit\(\)/);
  assert.doesNotMatch(preflight, /Start-Process[^\r\n]*-Wait\b/);
});

function run(script) {
  const command = `. '${core}'; ${script}`;
  return execFileSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', Buffer.from(command, 'utf16le').toString('base64')], {
    encoding: 'utf8',
  }).trim();
}

test('windows-distribution / upgrade preflight identifies only installed v2 product processes', () => {
  const result = run("$root='C:\\Users\\Reporter\\AppData\\Local'; @('C:\\Users\\Reporter\\AppData\\Local\\StockReporterAssistant\\app-1.0.13\\StockReporterAssistant.exe', 'C:\\Users\\Reporter\\AppData\\Local\\Packages\\OpenAI.Codex_abc\\LocalCache\\Local\\StockReporterAssistant\\app-1.0.11\\StockReporterAssistant.exe', 'F:\\SideProject\\stock-info-collector\\out\\StockReporterAssistant.exe', 'C:\\Users\\Reporter\\AppData\\Local\\OtherApp\\StockReporterAssistant.exe') | ForEach-Object { Test-ReporterInstalledExecutable $_ $root } | ConvertTo-Json -Compress");
  assert.deepEqual(JSON.parse(result), [true, true, false, false]);
});

test('windows-distribution / declining close does not stop a process or launch setup', () => {
  const result = run("$calls=[Collections.Generic.List[string]]::new(); $p=@([pscustomobject]@{ProcessId=42;ExecutablePath='C:\\Users\\Reporter\\AppData\\Local\\StockReporterAssistant\\app-1.0.13\\StockReporterAssistant.exe'}); $out=Invoke-ReporterInstallerPlan $p 'C:\\Users\\Reporter\\AppData\\Local' { param($count) $false } { param($id) $calls.Add('stop') } { param($id) $true } { $calls.Add('launch') }; [pscustomobject]@{outcome=$out;calls=@($calls)} | ConvertTo-Json -Compress");
  assert.deepEqual(JSON.parse(result), { outcome: 'cancelled', calls: [] });
});

test('windows-distribution / consent closes only verified instances before launching setup', () => {
  const result = run("$calls=[Collections.Generic.List[string]]::new(); $p=@([pscustomobject]@{ProcessId=42;ExecutablePath='C:\\Users\\Reporter\\AppData\\Local\\StockReporterAssistant\\app-1.0.13\\StockReporterAssistant.exe'},[pscustomobject]@{ProcessId=43;ExecutablePath='F:\\Other\\StockReporterAssistant.exe'}); $out=Invoke-ReporterInstallerPlan $p 'C:\\Users\\Reporter\\AppData\\Local' { param($count) $count -eq 1 } { param($pidToStop) $calls.Add('stop:'+ $pidToStop) } { param($pidToWait) $calls.Add('wait:'+ $pidToWait); $true } { $calls.Add('launch') }; [pscustomobject]@{outcome=$out;calls=@($calls)} | ConvertTo-Json -Compress");
  assert.deepEqual(JSON.parse(result), { outcome: 'installed', calls: ['stop:42', 'wait:42', 'launch'] });
});

test('windows-distribution / unknown product path aborts rather than installing through a stale singleton', () => {
  const result = run("$calls=[Collections.Generic.List[string]]::new(); $p=@([pscustomobject]@{ProcessId=42;ExecutablePath=$null}); $out=Invoke-ReporterInstallerPlan $p 'C:\\Users\\Reporter\\AppData\\Local' { $true } { $calls.Add('stop') } { $true } { $calls.Add('launch') }; [pscustomobject]@{outcome=$out;calls=@($calls)} | ConvertTo-Json -Compress");
  assert.deepEqual(JSON.parse(result), { outcome: 'unknown', calls: [] });
});

test('windows-distribution / a child that exits during shutdown does not abort the approved upgrade', () => {
  const result = run("$calls=[Collections.Generic.List[string]]::new(); $p=@([pscustomobject]@{ProcessId=42;ExecutablePath='C:\\Users\\Reporter\\AppData\\Local\\StockReporterAssistant\\app-1.0.13\\StockReporterAssistant.exe'},[pscustomobject]@{ProcessId=43;ExecutablePath='C:\\Users\\Reporter\\AppData\\Local\\StockReporterAssistant\\app-1.0.13\\StockReporterAssistant.exe'}); $out=Invoke-ReporterInstallerPlan $p 'C:\\Users\\Reporter\\AppData\\Local' { $true } { param($processId) if ($processId -eq 43) { throw 'already exited' }; $calls.Add('stop:'+ $processId) } { param($processId) $true } { $calls.Add('launch') }; [pscustomobject]@{outcome=$out;calls=@($calls)} | ConvertTo-Json -Compress");
  assert.deepEqual(JSON.parse(result), { outcome: 'installed', calls: ['stop:42', 'launch'] });
});

test('windows-distribution / an unclosable installed process prevents setup from starting', () => {
  const result = run("$calls=[Collections.Generic.List[string]]::new(); $p=@([pscustomobject]@{ProcessId=42;ExecutablePath='C:\\Users\\Reporter\\AppData\\Local\\StockReporterAssistant\\app-1.0.13\\StockReporterAssistant.exe'}); $out=Invoke-ReporterInstallerPlan $p 'C:\\Users\\Reporter\\AppData\\Local' { $true } { throw 'access denied' } { $false } { $calls.Add('launch') }; [pscustomobject]@{outcome=$out;calls=@($calls)} | ConvertTo-Json -Compress");
  assert.deepEqual(JSON.parse(result), { outcome: 'close_failed', calls: [] });
});

test('windows-distribution / first install proceeds without a close prompt', () => {
  const result = run("$calls=[Collections.Generic.List[string]]::new(); $out=Invoke-ReporterInstallerPlan @() 'C:\\Users\\Reporter\\AppData\\Local' { $calls.Add('prompt'); $false } { $calls.Add('stop') } { $true } { $calls.Add('launch') }; [pscustomobject]@{outcome=$out;calls=@($calls)} | ConvertTo-Json -Compress");
  assert.deepEqual(JSON.parse(result), { outcome: 'installed', calls: ['launch'] });
});
