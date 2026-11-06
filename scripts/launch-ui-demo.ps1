$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$repositoryRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $repositoryRoot
$url = 'http://127.0.0.1:5178'
try {
  $existing = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 2
  if ($existing.Content.Contains('name="reporter-ui-demo"')) {
    Start-Process $url
    exit 0
  }
} catch { }
$nodeRoot = Join-Path $env:LOCALAPPDATA 'mise\installs\node'
$npmPath = Get-ChildItem -LiteralPath $nodeRoot -Filter npm.cmd -File -Recurse -ErrorAction SilentlyContinue |
  Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
if (-not $npmPath) { $npmPath = (Get-Command npm.cmd -ErrorAction SilentlyContinue).Source }
if (-not $npmPath) { throw 'npm.cmd was not found. Install Node.js as documented in README.md.' }
$env:Path = (Split-Path -Parent $npmPath) + ';' + $env:Path
Write-Host 'Starting isolated UI demo. No backend, Google account or Windows schedule is used.'
Write-Host 'Keep this terminal open. Press Ctrl+C to stop the demo.'
& $npmPath run demo -- --open
exit $LASTEXITCODE
