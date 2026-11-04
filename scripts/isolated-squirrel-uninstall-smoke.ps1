$ErrorActionPreference = 'Stop'

if ($env:OS -ne 'Windows_NT') { throw 'Windows Task Scheduler is required.' }

$projectRoot = Split-Path -Parent $PSScriptRoot
$artifacts = Join-Path $projectRoot 'artifacts'
$metadata = Get-Content -LiteralPath (Join-Path $projectRoot 'package.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$packageDir = Join-Path $artifacts ('forge-out/' + $metadata.productName + '-win32-x64')
if (-not (Test-Path -LiteralPath (Join-Path $packageDir 'StockReporterAssistant.exe') -PathType Leaf)) {
  throw 'Packaged app is missing; run npm run build first.'
}

$stage = Join-Path $artifacts ('isolated-squirrel-uninstall-' + [guid]::NewGuid().ToString('N'))
$appDir = Join-Path $stage ('app-' + $metadata.version)
$updateExe = Join-Path $stage 'Update.exe'
$folderName = 'StockReporterAssistantSmoke-' + [guid]::NewGuid().ToString('N')
$folderPath = '\' + $folderName
$taskName = 'StockReporterAssistant-ConferenceDaily'
$taskArguments = '--processStart StockReporterAssistant.exe --process-start-args "--conference-scheduled"'

$service = New-Object -ComObject 'Schedule.Service'
$service.Connect()
$root = $service.GetFolder('\')
function Get-RootTaskXml([string]$Name) {
  try { return [string]$root.GetTask($Name).Xml }
  catch {
    $code = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$_.Exception.HResult), 0)
    if ($code -eq [Convert]::ToUInt32('80070002', 16) -or
        $code -eq [Convert]::ToUInt32('8004130F', 16)) { return $null }
    throw
  }
}
$rootV1Before = Get-RootTaskXml 'stock-info-collector-daily-work'
$rootAlternateV1Before = Get-RootTaskXml 'stock-info-collector'
$rootV2Before = Get-RootTaskXml $taskName
$folderCreated = $false
$launched = $null
$previousUserData = $env:REPORTER_USER_DATA_DIR
$previousTray = $env:REPORTER_TEST_TRAY
$previousLive = $env:REPORTER_TEST_LIVE_SOURCES

try {
  New-Item -ItemType Directory -Path $stage | Out-Null
  New-Item -ItemType Directory -Path $appDir | Out-Null
  foreach ($file in Get-ChildItem -LiteralPath $packageDir -Recurse -File) {
    if ($file.Name -match '^(?i:token[^\\]*\.json|credentials\.json|oauth-token\.enc)$') {
      throw 'Packaged app unexpectedly contains a user credential or token.'
    }
    if ($file.Name -eq 'google-oauth-client.json') { continue }
    $relative = $file.FullName.Substring($packageDir.Length).TrimStart('\')
    $destination = Join-Path $appDir $relative
    $destinationDirectory = Split-Path -Parent $destination
    if (-not (Test-Path -LiteralPath $destinationDirectory)) {
      New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
    }
    Copy-Item -LiteralPath $file.FullName -Destination $destination
  }
  Copy-Item -LiteralPath (Join-Path $env:SystemRoot 'System32/where.exe') -Destination $updateExe
  $gateway = Join-Path $appDir 'resources/conference-task.ps1'
  $source = [IO.File]::ReadAllText($gateway)
  $rootLookup = "GetFolder('\')"
  if ($source.Split([string[]]@($rootLookup), [StringSplitOptions]::None).Length -ne 2) {
    throw 'Expected exactly one Task Scheduler root lookup in the packaged gateway.'
  }
  $isolatedGateway = $source.Replace($rootLookup, "GetFolder('$folderPath')")
  [IO.File]::WriteAllText($gateway, $isolatedGateway, [Text.UTF8Encoding]::new($false))

  $root.CreateFolder($folderName) | Out-Null
  $folderCreated = $true
  $folder = $service.GetFolder($folderPath)
  $registration = & $gateway `
    -Action Register -TaskName $taskName -Executable $updateExe -WorkingDirectory $stage -Arguments $taskArguments
  if (($registration | Select-Object -Last 1 | ConvertFrom-Json).status -ne 'registered') {
    throw 'Isolated conference task registration failed.'
  }

  $userData = Join-Path $stage 'userData'
  New-Item -ItemType Directory -Path $userData | Out-Null
  $env:REPORTER_USER_DATA_DIR = $userData
  $env:REPORTER_TEST_TRAY = '0'
  Remove-Item Env:REPORTER_TEST_LIVE_SOURCES -ErrorAction SilentlyContinue
  $launched = Start-Process -FilePath (Join-Path $appDir 'StockReporterAssistant.exe') `
    -ArgumentList '--squirrel-uninstall' -PassThru -WindowStyle Hidden
  if (-not $launched.WaitForExit(30000)) { throw 'Packaged Squirrel uninstall event did not exit.' }
  if ($launched.ExitCode -ne 0) { throw "Packaged Squirrel uninstall exited $($launched.ExitCode)." }
  $taskStillExists = $true
  try { $folder.GetTask($taskName) | Out-Null }
  catch {
    $code = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$_.Exception.HResult), 0)
    if ($code -eq [Convert]::ToUInt32('80070002', 16) -or
        $code -eq [Convert]::ToUInt32('8004130F', 16)) { $taskStillExists = $false }
    else { throw }
  }
  if ($taskStillExists) {
    $taskDefinition = $folder.GetTask($taskName).Definition
    $taskAction = $taskDefinition.Actions.Item(1)
    $taskTrigger = $taskDefinition.Triggers.Item(1)
    $ownerSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $principalSid = try {
      (New-Object Security.Principal.NTAccount([string]$taskDefinition.Principal.UserId)).Translate(
        [Security.Principal.SecurityIdentifier]).Value
    } catch { '' }
    $checks = [ordered]@{
      actionCount = $taskDefinition.Actions.Count
      triggerCount = $taskDefinition.Triggers.Count
      dailyTrigger = ($taskTrigger.Type -eq 2)
      ignoreNew = ($taskDefinition.Settings.MultipleInstances -eq 2)
      interactiveUser = ($taskDefinition.Principal.LogonType -eq 3)
      ownerPath = [string]::Equals([IO.Path]::GetFullPath([string]$taskAction.Path),
        [IO.Path]::GetFullPath($updateExe), [StringComparison]::OrdinalIgnoreCase)
      workDirectory = [string]::Equals([IO.Path]::GetFullPath([string]$taskAction.WorkingDirectory),
        [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($updateExe)), [StringComparison]::OrdinalIgnoreCase)
      arguments = ([string]$taskAction.Arguments -eq $taskArguments)
      principal = ($principalSid -eq $ownerSid)
      triggerTime = ([string]$taskTrigger.StartBoundary -match '(T17:45:00|T19:00:00)([+-][0-9]{2}:[0-9]{2})?$')
    }
    $logPath = Join-Path $userData 'logs/application.log'
    $events = if (Test-Path -LiteralPath $logPath) {
      @(Get-Content -LiteralPath $logPath -Encoding UTF8 | ForEach-Object {
        try {
          $entry = $_ | ConvertFrom-Json
          if ($entry.event -like 'squirrel-*') { "$($entry.event):$($entry.outcome):$($entry.error.message)" }
        } catch { }
      }) -join '; '
    } else { 'no-app-log' }
    throw "Packaged Squirrel uninstall did not remove its isolated owned task. Checks: $($checks | ConvertTo-Json -Compress). Events: $events"
  }
  if ((Get-RootTaskXml 'stock-info-collector-daily-work') -cne $rootV1Before -or
      (Get-RootTaskXml 'stock-info-collector') -cne $rootAlternateV1Before -or
      (Get-RootTaskXml $taskName) -cne $rootV2Before) {
    throw 'A formal conference task changed during packaged uninstall smoke.'
  }
  Write-Output 'PACKAGED_SQUIRREL_UNINSTALL_ISOLATED_OK'
} finally {
  if ($launched -and -not $launched.HasExited) { Stop-Process -Id $launched.Id -Force -ErrorAction SilentlyContinue }
  if ($null -eq $previousUserData) { Remove-Item Env:REPORTER_USER_DATA_DIR -ErrorAction SilentlyContinue }
  else { $env:REPORTER_USER_DATA_DIR = $previousUserData }
  if ($null -eq $previousTray) { Remove-Item Env:REPORTER_TEST_TRAY -ErrorAction SilentlyContinue }
  else { $env:REPORTER_TEST_TRAY = $previousTray }
  if ($null -eq $previousLive) { Remove-Item Env:REPORTER_TEST_LIVE_SOURCES -ErrorAction SilentlyContinue }
  else { $env:REPORTER_TEST_LIVE_SOURCES = $previousLive }
  if ($folderCreated) {
    $folder = $service.GetFolder($folderPath)
    try {
      $remaining = $folder.GetTask($taskName)
      $action = $remaining.Definition.Actions.Item(1)
      if (-not [string]::Equals([string]$action.Path, $updateExe, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Unexpected task in isolated folder; preserving it for inspection.'
      }
      $folder.DeleteTask($taskName, 0)
    } catch [IO.FileNotFoundException] {
    } catch [Runtime.InteropServices.COMException] {
      $code = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$_.Exception.HResult), 0)
      if ($code -ne [Convert]::ToUInt32('80070002', 16) -and
          $code -ne [Convert]::ToUInt32('8004130F', 16)) { throw }
    }
    $root.DeleteFolder($folderName, 0)
  }
  $resolvedArtifacts = [IO.Path]::GetFullPath($artifacts).TrimEnd('\') + '\'
  $resolvedStage = [IO.Path]::GetFullPath($stage)
  if (-not $resolvedStage.StartsWith($resolvedArtifacts, [StringComparison]::OrdinalIgnoreCase) -or
      -not [IO.Path]::GetFileName($resolvedStage).StartsWith('isolated-squirrel-uninstall-')) {
    throw 'Isolated stage escaped the expected artifacts directory.'
  }
  for ($attempt = 0; $attempt -lt 15 -and (Test-Path -LiteralPath $resolvedStage); $attempt++) {
    try { Remove-Item -LiteralPath $resolvedStage -Recurse -Force -ErrorAction Stop }
    catch {
      if ($attempt -eq 14) { throw }
      Start-Sleep -Milliseconds 250
    }
  }
}
