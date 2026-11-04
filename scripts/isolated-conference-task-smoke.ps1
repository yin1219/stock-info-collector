$ErrorActionPreference = 'Stop'

if ($env:OS -ne 'Windows_NT') { throw 'Windows Task Scheduler is required.' }

$taskName = 'StockReporterAssistant-ConferenceDaily'
$folderName = 'StockReporterAssistantSmoke-' + [guid]::NewGuid().ToString('N')
$folderPath = '\' + $folderName
$stage = Join-Path ([IO.Path]::GetTempPath()) ('stock-reporter-task-smoke-' + [guid]::NewGuid().ToString('N'))
$gateway = Join-Path $PSScriptRoot 'conference-task.ps1'
$source = [IO.File]::ReadAllText($gateway)
$rootLookup = "GetFolder('\')"
if ($source.Split([string[]]@($rootLookup), [StringSplitOptions]::None).Length -ne 2) {
  throw 'Expected exactly one Task Scheduler root lookup in the production gateway.'
}
$isolatedGateway = $source.Replace($rootLookup, "GetFolder('$folderPath')")
if ($isolatedGateway.Contains($rootLookup)) { throw 'Gateway isolation failed.' }

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
$alternateLegacyCreated = $false
$previousGateway = $env:REPORTER_TEST_TASK_GATEWAY
$previousAction = $env:REPORTER_TEST_TASK_ACTION
$previousExecutable = $env:REPORTER_TEST_TASK_EXECUTABLE
$previousDirectory = $env:REPORTER_TEST_TASK_DIRECTORY
$previousBackupDirectory = $env:REPORTER_TEST_TASK_BACKUP_DIRECTORY

try {
  New-Item -ItemType Directory -Path $stage | Out-Null
  $patchedPath = Join-Path $stage 'conference-task-isolated.ps1'
  $invokePath = Join-Path $stage 'invoke.ps1'
  [IO.File]::WriteAllText($patchedPath, $isolatedGateway, [Text.UTF8Encoding]::new($false))
  [IO.File]::WriteAllText($invokePath, @'
& $env:REPORTER_TEST_TASK_GATEWAY -Action $env:REPORTER_TEST_TASK_ACTION `
  -TaskName 'StockReporterAssistant-ConferenceDaily' `
  -Executable $env:REPORTER_TEST_TASK_EXECUTABLE `
  -WorkingDirectory $env:REPORTER_TEST_TASK_DIRECTORY `
  -BackupDirectory $env:REPORTER_TEST_TASK_BACKUP_DIRECTORY `
  -Arguments '--processStart StockReporterAssistant.exe --process-start-args "--conference-scheduled"'
'@, [Text.UTF8Encoding]::new($false))
  $env:REPORTER_TEST_TASK_GATEWAY = $patchedPath
  $env:REPORTER_TEST_TASK_EXECUTABLE = Join-Path $stage 'Update.exe'
  $env:REPORTER_TEST_TASK_DIRECTORY = $stage
  $env:REPORTER_TEST_TASK_BACKUP_DIRECTORY = Join-Path $stage 'backups'

  $root.CreateFolder($folderName) | Out-Null
  $folderCreated = $true
  $folder = $service.GetFolder($folderPath)

  function Invoke-IsolatedGateway([string]$Action) {
    $env:REPORTER_TEST_TASK_ACTION = $Action
    $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $invokePath
    if ($LASTEXITCODE -ne 0) { throw "Isolated gateway $Action failed: $($output -join ' ')" }
    return ($output | Select-Object -Last 1 | ConvertFrom-Json)
  }

  $registered = Invoke-IsolatedGateway 'Register'
  if ($registered.status -ne 'registered') { throw 'Temporary v2 task was not registered.' }
  $inspected = Invoke-IsolatedGateway 'Inspect'
  if ($inspected.status -ne 'found' -or
      $inspected.task.dailyAtLocal -ne '19:00' -or
      -not $inspected.task.enabled -or
      $inspected.task.multipleInstances -ne 'IgnoreNew' -or
      -not [string]::Equals($inspected.task.execute, $env:REPORTER_TEST_TASK_EXECUTABLE, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Temporary v2 task did not match its expected action and daily trigger.'
  }
  $oldDefinition = $folder.GetTask($taskName).Definition
  $oldDefinition.Triggers.Item(1).StartBoundary = (Get-Date -Format 'yyyy-MM-dd') + 'T17:45:00+08:00'
  $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
  $folder.RegisterTaskDefinition($taskName, $oldDefinition, 6, $currentUser, $null, 3) | Out-Null
  $oldXml = [string]$folder.GetTask($taskName).Xml
  $migrated = Invoke-IsolatedGateway 'MigrateTime'
  if ($migrated.status -ne 'migrated' -or
      -not [IO.File]::Exists([string]$migrated.backup) -or
      [IO.File]::ReadAllText([string]$migrated.backup) -cne $oldXml -or
      (Invoke-IsolatedGateway 'Inspect').task.dailyAtLocal -ne '19:00') {
    throw 'Temporary v2 task did not migrate from 17:45 to 19:00 with a verified backup.'
  }
  $removed = Invoke-IsolatedGateway 'Remove'
  if ($removed.status -ne 'removed') { throw 'Temporary v2 task was not removed by its owner check.' }
  $missing = Invoke-IsolatedGateway 'Inspect'
  if ($missing.status -ne 'missing') { throw 'Temporary v2 task remained after removal.' }
  if ((Get-RootTaskXml 'stock-info-collector-daily-work') -cne $rootV1Before -or
      (Get-RootTaskXml 'stock-info-collector') -cne $rootAlternateV1Before -or
      (Get-RootTaskXml $taskName) -cne $rootV2Before) {
    throw 'A root-level conference task changed during the isolated smoke test.'
  }
  Write-Output 'CONFERENCE_TASK_WINDOWS_ISOLATED_OK'

  $foreignExecutable = Join-Path $stage 'Foreign.exe'
  $foreignDefinition = $service.NewTask(0)
  $foreignDefinition.Settings.MultipleInstances = 2
  $foreignDefinition.Principal.LogonType = 3
  $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
  $foreignDefinition.Principal.UserId = $currentUser
  $foreignTrigger = $foreignDefinition.Triggers.Create(2)
  $foreignTrigger.StartBoundary = (Get-Date -Format 'yyyy-MM-dd') + 'T19:00:00'
  $foreignTrigger.DaysInterval = 1
  $foreignAction = $foreignDefinition.Actions.Create(0)
  $foreignAction.Path = $foreignExecutable
  $foreignAction.WorkingDirectory = $stage
  $foreignAction.Arguments = '--not-stock-reporter'
  $folder.RegisterTaskDefinition($taskName, $foreignDefinition, 2, $currentUser, $null, 3) | Out-Null
  $foreignRemoval = Invoke-IsolatedGateway 'Remove'
  $foreignMigrationBlocked = $false
  try { Invoke-IsolatedGateway 'MigrateTime' | Out-Null } catch { $foreignMigrationBlocked = $true }
  if ($foreignRemoval.status -ne 'skipped' -or
      -not $foreignMigrationBlocked -or
      -not [string]::Equals([string]$folder.GetTask($taskName).Definition.Actions.Item(1).Path,
        $foreignExecutable, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'A same-name foreign task was removed or changed.'
  }
  $folder.DeleteTask($taskName, 0)
  if ((Get-RootTaskXml 'stock-info-collector-daily-work') -cne $rootV1Before -or
      (Get-RootTaskXml 'stock-info-collector') -cne $rootAlternateV1Before -or
      (Get-RootTaskXml $taskName) -cne $rootV2Before) {
    throw 'A root-level conference task changed during foreign-task refusal.'
  }
  Write-Output 'CONFERENCE_TASK_FOREIGN_PRESERVED_OK'

  $alternateName = 'stock-info-collector'
  $alternateExe = Join-Path $stage 'stock-info-collector.exe'
  $alternateDefinition = $service.NewTask(0)
  $alternateDefinition.Settings.MultipleInstances = 2
  $alternateDefinition.Principal.LogonType = 3
  $alternateDefinition.Principal.UserId = $currentUser
  $alternateTrigger = $alternateDefinition.Triggers.Create(2)
  $alternateTrigger.StartBoundary = (Get-Date).AddDays(1).ToString('yyyy-MM-dd') + 'T18:15:00'
  $alternateTrigger.DaysInterval = 1
  $alternateAction = $alternateDefinition.Actions.Create(0)
  $alternateAction.Path = $alternateExe
  $alternateAction.WorkingDirectory = $stage + '\'
  $folder.RegisterTaskDefinition($alternateName, $alternateDefinition, 2, $currentUser, $null, 3) | Out-Null
  $alternateLegacyCreated = $true
  $secondWriterBlocked = $false
  try { Invoke-IsolatedGateway 'Register' | Out-Null } catch { $secondWriterBlocked = $true }
  if (-not $secondWriterBlocked -or (Invoke-IsolatedGateway 'Inspect').status -ne 'missing') {
    throw 'An active older v1 task did not prevent v2 registration.'
  }
  $guardedDefinition = $service.NewTask(0)
  $guardedDefinition.Settings.MultipleInstances = 2
  $guardedDefinition.Principal.LogonType = 3
  $guardedDefinition.Principal.UserId = $currentUser
  $guardedTrigger = $guardedDefinition.Triggers.Create(2)
  $guardedTrigger.StartBoundary = (Get-Date).AddDays(1).ToString('yyyy-MM-dd') + 'T17:45:00+08:00'
  $guardedTrigger.DaysInterval = 1
  $guardedAction = $guardedDefinition.Actions.Create(0)
  $guardedAction.Path = $env:REPORTER_TEST_TASK_EXECUTABLE
  $guardedAction.WorkingDirectory = $stage
  $guardedAction.Arguments = '--processStart StockReporterAssistant.exe --process-start-args "--conference-scheduled"'
  $folder.RegisterTaskDefinition($taskName, $guardedDefinition, 2, $currentUser, $null, 3) | Out-Null
  $guardedXml = [string]$folder.GetTask($taskName).Xml
  $migrationBlocked = $false
  try { Invoke-IsolatedGateway 'MigrateTime' | Out-Null } catch { $migrationBlocked = $true }
  if (-not $migrationBlocked -or [string]$folder.GetTask($taskName).Xml -cne $guardedXml) {
    throw 'An active older v1 task did not prevent v2 time migration.'
  }
  $folder.GetTask($alternateName).Enabled = $false
  $migratedAfterDisable = Invoke-IsolatedGateway 'MigrateTime'
  if ($migratedAfterDisable.status -ne 'migrated' -or
      (Invoke-IsolatedGateway 'Inspect').task.dailyAtLocal -ne '19:00') {
    throw 'A disabled older v1 task with a trailing working-directory separator blocked a safe v2 migration.'
  }
  Write-Output 'CONFERENCE_TASK_DISABLED_ALTERNATE_TRAILING_PATH_OK'
  $folder.DeleteTask($taskName, 0)
  Write-Output 'CONFERENCE_TASK_ALTERNATE_V1_MIGRATION_BLOCKED_OK'
  $folder.DeleteTask($alternateName, 0)
  $alternateLegacyCreated = $false
  if ((Get-RootTaskXml 'stock-info-collector-daily-work') -cne $rootV1Before -or
      (Get-RootTaskXml 'stock-info-collector') -cne $rootAlternateV1Before -or
      (Get-RootTaskXml $taskName) -cne $rootV2Before) {
    throw 'A root-level conference task changed during older-v1 refusal.'
  }
  Write-Output 'CONFERENCE_TASK_ALTERNATE_V1_BLOCKED_OK'
} finally {
  if ($folderCreated) {
    if ($alternateLegacyCreated) {
      $alternate = $folder.GetTask('stock-info-collector')
      if (-not [string]::Equals([string]$alternate.Definition.Actions.Item(1).Path,
          (Join-Path $stage 'stock-info-collector.exe'), [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Unexpected older-v1 task in isolated folder; preserving it for inspection.'
      }
      $folder.DeleteTask('stock-info-collector', 0)
    }
    try {
      $remaining = $folder.GetTask($taskName)
      $action = $remaining.Definition.Actions.Item(1)
      if ([string]::Equals([string]$action.Path, (Join-Path $stage 'Update.exe'), [StringComparison]::OrdinalIgnoreCase) -or
          [string]::Equals([string]$action.Path, (Join-Path $stage 'Foreign.exe'), [StringComparison]::OrdinalIgnoreCase)) {
        $folder.DeleteTask($taskName, 0)
      } else { throw 'Unexpected task in isolated smoke folder; preserving it for inspection.' }
    } catch [IO.FileNotFoundException] {} catch [Runtime.InteropServices.COMException] {
      $code = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$_.Exception.HResult), 0)
      if ($code -ne [Convert]::ToUInt32('80070002', 16) -and $code -ne [Convert]::ToUInt32('8004130F', 16)) { throw }
    }
    $root.DeleteFolder($folderName, 0)
  }
  if ($null -eq $previousGateway) { Remove-Item Env:REPORTER_TEST_TASK_GATEWAY -ErrorAction SilentlyContinue }
  else { $env:REPORTER_TEST_TASK_GATEWAY = $previousGateway }
  if ($null -eq $previousAction) { Remove-Item Env:REPORTER_TEST_TASK_ACTION -ErrorAction SilentlyContinue }
  else { $env:REPORTER_TEST_TASK_ACTION = $previousAction }
  if ($null -eq $previousExecutable) { Remove-Item Env:REPORTER_TEST_TASK_EXECUTABLE -ErrorAction SilentlyContinue }
  else { $env:REPORTER_TEST_TASK_EXECUTABLE = $previousExecutable }
  if ($null -eq $previousDirectory) { Remove-Item Env:REPORTER_TEST_TASK_DIRECTORY -ErrorAction SilentlyContinue }
  else { $env:REPORTER_TEST_TASK_DIRECTORY = $previousDirectory }
  if ($null -eq $previousBackupDirectory) { Remove-Item Env:REPORTER_TEST_TASK_BACKUP_DIRECTORY -ErrorAction SilentlyContinue }
  else { $env:REPORTER_TEST_TASK_BACKUP_DIRECTORY = $previousBackupDirectory }
  Remove-Item -LiteralPath (Join-Path $stage 'conference-task-isolated.ps1') -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath (Join-Path $stage 'invoke.ps1') -Force -ErrorAction SilentlyContinue
  if ($migrated -and [string]$migrated.backup -and
      [string]::Equals([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath([string]$migrated.backup)),
        [IO.Path]::GetFullPath((Join-Path $stage 'backups')), [StringComparison]::OrdinalIgnoreCase)) {
    Remove-Item -LiteralPath ([string]$migrated.backup) -Force -ErrorAction SilentlyContinue
  }
  if ($migratedAfterDisable -and [string]$migratedAfterDisable.backup -and
      [string]::Equals([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath([string]$migratedAfterDisable.backup)),
        [IO.Path]::GetFullPath((Join-Path $stage 'backups')), [StringComparison]::OrdinalIgnoreCase)) {
    Remove-Item -LiteralPath ([string]$migratedAfterDisable.backup) -Force -ErrorAction SilentlyContinue
  }
  Remove-Item -LiteralPath (Join-Path $stage 'backups') -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $stage -ErrorAction SilentlyContinue
}
