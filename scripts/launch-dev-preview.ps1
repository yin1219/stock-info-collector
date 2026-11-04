param([switch]$EnableGoogleOAuth, [switch]$ValidateOnly)

if ($ValidateOnly) {
  Write-Output 'DEV_PREVIEW_LAUNCHER_OK'
  exit 0
}

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$nodeRoot = Join-Path $env:LOCALAPPDATA 'mise\installs\node'
$npmPath = Get-ChildItem -LiteralPath $nodeRoot -Filter npm.cmd -File -Recurse -ErrorAction SilentlyContinue |
  Sort-Object FullName -Descending |
  Select-Object -First 1 -ExpandProperty FullName
if (-not $npmPath) {
  $npmCommand = Get-Command npm.cmd -ErrorAction SilentlyContinue
  $npmPath = $npmCommand.Source
}
if (-not $npmPath) { throw 'npm.cmd was not found. Install the Node.js version documented in README.md.' }

$previewData = Join-Path $env:TEMP ('stock-reporter-dev-preview-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $previewData | Out-Null
$env:Path = (Split-Path -Parent $npmPath) + ';' + $env:Path
$env:REPORTER_USER_DATA_DIR = $previewData
$env:REPORTER_TEST_TRAY = '1'
Remove-Item Env:REPORTER_TEST_LIVE_SOURCES -ErrorAction SilentlyContinue
Remove-Item Env:REPORTER_TEST_HTTP_PROXY -ErrorAction SilentlyContinue
if ($EnableGoogleOAuth) { $env:REPORTER_TEST_LIVE_OAUTH = '1' }
else { Remove-Item Env:REPORTER_TEST_LIVE_OAUTH -ErrorAction SilentlyContinue }

Write-Host 'Starting Stock Reporter Assistant v2 isolated preview. Use the tray menu to quit.'
Write-Host "Isolated data and logs: $previewData"
Set-Location -LiteralPath $repositoryRoot
& $npmPath start
exit $LASTEXITCODE
