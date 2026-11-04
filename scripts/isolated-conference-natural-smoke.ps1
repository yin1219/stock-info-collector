param(
  [string]$PackagedExecutable = ''
)

$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Windows Task Scheduler is required.' }
if (-not $PackagedExecutable) {
  $projectRoot = Split-Path -Parent $PSScriptRoot
  $metadata = Get-Content -LiteralPath (Join-Path $projectRoot 'package.json') -Raw -Encoding UTF8 | ConvertFrom-Json
  $PackagedExecutable = Join-Path $projectRoot ('artifacts/forge-out/' + $metadata.productName + '-win32-x64/StockReporterAssistant.exe')
}
if (-not (Test-Path -LiteralPath $PackagedExecutable -PathType Leaf)) { throw 'Packaged app is missing.' }

$stage = Join-Path ([IO.Path]::GetTempPath()) ('reporter-scheduled-smoke-' + [guid]::NewGuid().ToString('N'))
$userData = Join-Path $stage 'userData'
$runner = Join-Path $stage 'runner.ps1'
$resultPath = Join-Path $stage 'result.json'
$folderName = 'StockReporterNaturalSmoke-' + [guid]::NewGuid().ToString('N')
$taskName = 'IsolatedConferenceHeadless'
$service = New-Object -ComObject 'Schedule.Service'
$service.Connect()
$root = $service.GetFolder('\')
$folderCreated = $false

function Get-RootXml([string]$Name) {
  try { return [string]$root.GetTask($Name).Xml }
  catch {
    $code = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$_.Exception.HResult), 0)
    if ($code -eq [Convert]::ToUInt32('80070002', 16) -or
        $code -eq [Convert]::ToUInt32('8004130F', 16)) { return $null }
    throw
  }
}

$v1Before = Get-RootXml 'stock-info-collector-daily-work'
$alternateV1Before = Get-RootXml 'stock-info-collector'
$v2Before = Get-RootXml 'StockReporterAssistant-ConferenceDaily'
try {
  New-Item -ItemType Directory -Path $stage | Out-Null
  [IO.File]::WriteAllText($runner, @'
param([string]$Executable, [string]$UserData, [string]$ResultPath)
$ErrorActionPreference = 'Stop'
$env:REPORTER_USER_DATA_DIR = $UserData
Remove-Item Env:REPORTER_TEST_LIVE_SOURCES -ErrorAction SilentlyContinue
Remove-Item Env:REPORTER_TEST_LIVE_OAUTH -ErrorAction SilentlyContinue
Remove-Item Env:REPORTER_TEST_HTTP_PROXY -ErrorAction SilentlyContinue
$child = Start-Process -FilePath $Executable -ArgumentList '--conference-scheduled' -Wait -PassThru -WindowStyle Hidden
@{ exitCode = $child.ExitCode } | ConvertTo-Json -Compress | Set-Content -LiteralPath $ResultPath -Encoding ASCII
'@, [Text.UTF8Encoding]::new($false))

  $root.CreateFolder($folderName) | Out-Null
  $folderCreated = $true
  $folder = $service.GetFolder(('\' + $folderName))
  $definition = $service.NewTask(0)
  $definition.Settings.Enabled = $true
  $definition.Settings.MultipleInstances = 2
  $definition.Settings.StartWhenAvailable = $true
  $definition.Principal.LogonType = 3
  $userId = [Security.Principal.WindowsIdentity]::GetCurrent().Name
  $definition.Principal.UserId = $userId
  $trigger = $definition.Triggers.Create(2)
  $trigger.StartBoundary = (Get-Date).AddMinutes(1).ToString('yyyy-MM-ddTHH:mm:ss')
  $trigger.DaysInterval = 1
  $action = $definition.Actions.Create(0)
  $action.Path = (Join-Path $PSHOME 'powershell.exe')
  $action.WorkingDirectory = $stage
  $action.Arguments = '-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $runner + '" -Executable "' + $PackagedExecutable + '" -UserData "' + $userData + '" -ResultPath "' + $resultPath + '"'
  $task = $folder.RegisterTaskDefinition($taskName, $definition, 2, $userId, $null, 3)
  $deadline = (Get-Date).AddMinutes(3)
  while ((Get-Date) -lt $deadline -and -not (Test-Path -LiteralPath $resultPath)) {
    Start-Sleep -Seconds 2
  }
  if (-not (Test-Path -LiteralPath $resultPath)) { throw 'The isolated natural trigger did not finish before the deadline.' }
  $result = Get-Content -LiteralPath $resultPath -Raw -Encoding ASCII | ConvertFrom-Json
  $task = $folder.GetTask($taskName)
  if ([int]$result.exitCode -ne 0 -or [int64]$task.LastTaskResult -ne 0 -or
      $task.Definition.Triggers.Count -ne 1 -or
      [int]$task.Definition.Triggers.Item(1).Type -ne 2 -or
      [int]$task.Definition.Triggers.Item(1).DaysInterval -ne 1) {
    throw 'The isolated packaged headless task returned a nonzero result.'
  }
  $logPath = Join-Path $userData 'logs/application.log'
  if (-not (Test-Path -LiteralPath $logPath) -or
      -not ([IO.File]::ReadAllText($logPath).Contains('"event":"conference-scheduled-not-ready"'))) {
    throw 'The isolated scheduled app did not record the expected fail-closed result.'
  }
  if ((Get-RootXml 'stock-info-collector-daily-work') -cne $v1Before -or
      (Get-RootXml 'stock-info-collector') -cne $alternateV1Before -or
      (Get-RootXml 'StockReporterAssistant-ConferenceDaily') -cne $v2Before) {
    throw 'A formal root-level conference task changed.'
  }
  Write-Output ('CONFERENCE_DAILY_NATURAL_ISOLATED_OK trigger=' + $trigger.StartBoundary + ' triggerType=Daily exit=0 formal-tasks-unchanged')
} finally {
  if ($folderCreated) {
    try {
      $folder = $service.GetFolder(('\' + $folderName))
      $task = $folder.GetTask($taskName)
      if ([string]$task.Definition.Actions.Item(1).Path -ne (Join-Path $PSHOME 'powershell.exe') -or
          [string]$task.Definition.Actions.Item(1).WorkingDirectory -ne $stage) {
        throw 'Unexpected test task action; preserving the task for inspection.'
      }
      if ([int]$task.State -eq 4) { $task.Stop(0) }
      $folder.DeleteTask($taskName, 0)
    } catch [IO.FileNotFoundException] {} catch [Runtime.InteropServices.COMException] {
      $code = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$_.Exception.HResult), 0)
      if ($code -ne [Convert]::ToUInt32('80070002', 16) -and
          $code -ne [Convert]::ToUInt32('8004130F', 16)) { throw }
    }
    $root.DeleteFolder($folderName, 0)
  }
  $resolvedStage = [IO.Path]::GetFullPath($stage)
  $resolvedTemp = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
  if (-not $resolvedStage.StartsWith($resolvedTemp, [StringComparison]::OrdinalIgnoreCase) -or
      [IO.Path]::GetFileName($resolvedStage) -notlike 'reporter-scheduled-smoke-*') {
    throw 'Refusing to delete a stage outside the intended temporary directory.'
  }
  if (Test-Path -LiteralPath $resolvedStage) { Remove-Item -LiteralPath $resolvedStage -Recurse -Force }
}
