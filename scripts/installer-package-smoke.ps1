$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$metadata = Get-Content -LiteralPath (Join-Path $projectRoot 'package.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$makeDir = Join-Path $projectRoot 'artifacts/forge-out/make/squirrel.windows/x64'
$setup = Join-Path $makeDir 'StockReporterAssistantSetup.exe'
$package = Join-Path $makeDir ("StockReporterAssistant-{0}-full.nupkg" -f $metadata.version)

if (-not (Test-Path -LiteralPath $setup -PathType Leaf)) { throw 'Windows setup executable is missing; run npm run make first.' }
if (-not (Test-Path -LiteralPath $package -PathType Leaf)) { throw 'Matching full.nupkg is missing; run npm run make first.' }
$innerSetup = Join-Path $makeDir 'StockReporterAssistantSquirrelSetup.exe'
if (-not (Test-Path -LiteralPath $innerSetup -PathType Leaf)) { throw 'Prompted setup is missing its Squirrel source package.' }

function Get-ReporterSha256 {
  param([string]$Path)
  $stream = [IO.File]::OpenRead($Path)
  $sha256 = [Security.Cryptography.SHA256]::Create()
  try { [BitConverter]::ToString($sha256.ComputeHash($stream)).Replace('-', '') }
  finally {
    $sha256.Dispose()
    $stream.Dispose()
  }
}

$extractDir = Join-Path $makeDir ('installer-smoke-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $extractDir | Out-Null
try {
  $process = Start-Process -FilePath $setup -ArgumentList @('/C', ('/T:' + $extractDir)) -Wait -PassThru -WindowStyle Hidden
  if ($process.ExitCode -ne 0) { throw 'Prompted setup could not be extracted without installing.' }
  $embedded = Join-Path $extractDir 'StockReporterAssistantSquirrelSetup.exe'
  if (-not (Test-Path -LiteralPath $embedded -PathType Leaf)) { throw 'Prompted setup does not contain Squirrel setup.' }
  if ((Get-ReporterSha256 $embedded) -ne (Get-ReporterSha256 $innerSetup)) {
    throw 'Prompted setup contains a different Squirrel package.'
  }
  foreach ($name in 'installer-preflight.ps1', 'installer-preflight-core.ps1') {
    $embeddedScript = Join-Path $extractDir $name
    if (-not (Test-Path -LiteralPath $embeddedScript -PathType Leaf)) {
      throw "Prompted setup is missing $name."
    }
    $sourceScript = Join-Path $PSScriptRoot $name
    $sourceText = Get-Content -LiteralPath $sourceScript -Raw -Encoding UTF8
    $embeddedText = (Get-Content -LiteralPath $embeddedScript -Raw -Encoding UTF8).TrimStart([char]0xFEFF)
    if ($embeddedText -cne $sourceText) { throw "Prompted setup has outdated $name." }
  }
} finally {
  $resolvedMakeDir = [IO.Path]::GetFullPath($makeDir).TrimEnd('\') + '\'
  $resolvedExtractDir = [IO.Path]::GetFullPath($extractDir)
  if (-not $resolvedExtractDir.StartsWith($resolvedMakeDir, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Installer smoke extraction escaped the expected output directory.'
  }
  Remove-Item -LiteralPath $resolvedExtractDir -Recurse -Force
}

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
