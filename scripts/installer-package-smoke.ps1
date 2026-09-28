$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$metadata = Get-Content -LiteralPath (Join-Path $projectRoot 'package.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$makeDir = Join-Path $projectRoot 'artifacts/forge-out/make/squirrel.windows/x64'
$setup = Join-Path $makeDir 'StockReporterAssistantSetup.exe'
$package = Join-Path $makeDir ("StockReporterAssistant-{0}-full.nupkg" -f $metadata.version)

if (-not (Test-Path -LiteralPath $setup -PathType Leaf)) { throw 'Windows setup executable is missing; run npm run make first.' }
if (-not (Test-Path -LiteralPath $package -PathType Leaf)) { throw 'Matching full.nupkg is missing; run npm run make first.' }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($package)
try {
  $entries = @($archive.Entries | ForEach-Object { $_.FullName })
  if ($entries -cnotcontains 'lib/net45/StockReporterAssistant.exe') {
    throw 'Installer does not contain the expected ASCII executable name.'
  }
  if ($entries -match '(?i)(^|/)(token[^/]*\.json|credentials\.json|oauth-token\.enc)$') {
    throw 'Installer contains a user credential or token file.'
  }
  if ($env:REPORTER_EXPECT_OAUTH_CLIENT -eq '1' -and $entries -cnotcontains 'lib/net45/resources/google-oauth-client.json') {
    throw 'Installer does not contain the authorized app OAuth client.'
  }
} finally {
  $archive.Dispose()
}

Write-Output "Squirrel installer archive smoke passed for version $($metadata.version)."
