param(
  [ValidateSet('Inspect', 'Register', 'Remove', 'MigrateTime')][string]$Action = 'Inspect',
  [string]$TaskName,
  [string]$Executable,
  [string]$WorkingDirectory,
  [string]$Arguments,
  [string]$BackupDirectory,
  [switch]$AllowDisabledLegacy,
  [switch]$ValidateOnly
)

function Test-MissingTaskException([Exception]$Failure) {
  $candidate = $Failure
  while ($null -ne $candidate) {
    $code = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$candidate.HResult), 0)
    if ($code -eq [Convert]::ToUInt32('80070002', 16) -or
        $code -eq [Convert]::ToUInt32('8004130F', 16)) { return $true }
    $candidate = $candidate.InnerException
  }
  return $false
}

function Test-CurrentUserPrincipal([string]$PrincipalId) {
  if ([string]::IsNullOrWhiteSpace($PrincipalId)) { return $false }
  try {
    $currentSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $principalSid = if ($PrincipalId -match '^S-1-') {
      (New-Object Security.Principal.SecurityIdentifier($PrincipalId)).Value
    } else {
      (New-Object Security.Principal.NTAccount($PrincipalId)).Translate([Security.Principal.SecurityIdentifier]).Value
    }
    return [string]::Equals($principalSid, $currentSid, [StringComparison]::OrdinalIgnoreCase)
  } catch {
    return $false
  }
}

function Assert-LegacyTaskAbsent($Folder) {
  foreach ($legacyName in @('stock-info-collector-daily-work', 'stock-info-collector')) {
  $legacyTaskMissing = $false
  $legacyTask = $null
  try {
    $legacyTask = $Folder.GetTask($legacyName)
  } catch {
    if (Test-MissingTaskException $_.Exception) { $legacyTaskMissing = $true }
    else { throw }
  }
  if ($legacyTaskMissing) { continue }
  if (-not $AllowDisabledLegacy) { throw 'A legacy conference task exists; v2 task was not registered.' }
  if ([bool]$legacyTask.Enabled -or [int]$legacyTask.State -ne 1) {
    throw 'The legacy conference task is not confirmed disabled and stopped.'
  }
  $definition = $legacyTask.Definition
  if ($definition.Actions.Count -ne 1 -or $definition.Triggers.Count -ne 1 -or
      $definition.Triggers.Item(1).Type -ne 2) {
    throw 'The disabled legacy conference task is not recognized.'
  }
  $legacyAction = $definition.Actions.Item(1)
  $legacyPath = [IO.Path]::GetFullPath([string]$legacyAction.Path)
  $legacyDirectory = [IO.Path]::GetFullPath((Join-Path ([string]$legacyAction.WorkingDirectory) '.'))
  if ([IO.Path]::GetFileName($legacyPath) -ne 'stock-info-collector.exe' -or
      -not [string]::Equals([IO.Path]::GetDirectoryName($legacyPath), $legacyDirectory, [StringComparison]::OrdinalIgnoreCase) -or
      [string]$legacyAction.Arguments -ne '') {
    throw 'The disabled legacy conference task action is not recognized.'
  }
  }
}

if ($ValidateOnly) {
  if (-not (Test-MissingTaskException ([IO.FileNotFoundException]::new())) -or
      (Test-MissingTaskException ([UnauthorizedAccessException]::new()))) {
    throw 'Task Scheduler missing-task HRESULT classification failed.'
  }
  Write-Output 'CONFERENCE_TASK_GATEWAY_OK'
  Write-Output 'CONFERENCE_TASK_MISSING_CODES_OK'
  $missing = [pscustomobject]@{}
  $missing | Add-Member ScriptMethod GetTask { param($name) throw [IO.FileNotFoundException]::new() }
  $found = [pscustomobject]@{}
  $found | Add-Member ScriptMethod GetTask { param($name) [pscustomobject]@{ Name = $name } }
  $denied = [pscustomobject]@{}
  $denied | Add-Member ScriptMethod GetTask { param($name) throw [UnauthorizedAccessException]::new() }
  $missingAllowed = $false
  $foundBlocked = $false
  $deniedBlocked = $false
  try { Assert-LegacyTaskAbsent $missing; $missingAllowed = $true } catch {}
  try { Assert-LegacyTaskAbsent $found } catch { $foundBlocked = $true }
  try { Assert-LegacyTaskAbsent $denied } catch { $deniedBlocked = $true }
  if (-not $missingAllowed -or -not $foundBlocked -or -not $deniedBlocked) {
    throw 'Single-writer registration guard failed.'
  }
  Write-Output 'CONFERENCE_TASK_SINGLE_WRITER_OK'
  $fakeAction = [pscustomobject]@{ Path = 'C:\V1\stock-info-collector.exe'; WorkingDirectory = 'C:\V1'; Arguments = '' }
  $actions = [pscustomobject]@{ Count = 1; Entry = $fakeAction }
  $actions | Add-Member ScriptMethod Item { param($number) $this.Entry }
  $trigger = [pscustomobject]@{ Type = 2 }
  $triggers = [pscustomobject]@{ Count = 1; Entry = $trigger }
  $triggers | Add-Member ScriptMethod Item { param($number) $this.Entry }
  $definition = [pscustomobject]@{ Actions = $actions; Triggers = $triggers }
  $legacy = [pscustomobject]@{ Enabled = $false; State = 1; Definition = $definition }
  $retained = [pscustomobject]@{ Entry = $legacy }
  $retained | Add-Member ScriptMethod GetTask { param($name) $this.Entry }
  $AllowDisabledLegacy = $true
  Assert-LegacyTaskAbsent $retained
  $legacy.State = 4
  $runningBlocked = $false
  try { Assert-LegacyTaskAbsent $retained } catch { $runningBlocked = $true }
  $legacy.State = 1
  $legacy.Enabled = $true
  $enabledBlocked = $false
  try { Assert-LegacyTaskAbsent $retained } catch { $enabledBlocked = $true }
  if (-not $runningBlocked -or -not $enabledBlocked) { throw 'Disabled legacy transition guard failed.' }
  Write-Output 'CONFERENCE_TASK_DISABLED_LEGACY_OK'
  $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $shortName = $identity.Name.Split('\')[-1]
  if (-not (Test-CurrentUserPrincipal $identity.Name) -or
      -not (Test-CurrentUserPrincipal $identity.User.Value) -or
      -not (Test-CurrentUserPrincipal $shortName) -or
      (Test-CurrentUserPrincipal '') -or
      (Test-CurrentUserPrincipal 'S-1-5-18') -or
      (Test-CurrentUserPrincipal '__stock_reporter_unknown_account__')) {
    throw 'Task owner identity check failed.'
  }
  Write-Output 'CONFERENCE_TASK_OWNER_IDENTITY_OK'
  exit 0
}

$ErrorActionPreference = 'Stop'
if ($TaskName -notin @('stock-info-collector-daily-work', 'stock-info-collector', 'StockReporterAssistant-ConferenceDaily')) {
  throw 'Task name is not in the conference allowlist.'
}

$service = New-Object -ComObject 'Schedule.Service'
$service.Connect()
$folder = $service.GetFolder('\')

if ($Action -eq 'Inspect') {
  try {
    $task = $folder.GetTask($TaskName)
  } catch {
    if (Test-MissingTaskException $_.Exception) {
      @{ status = 'missing' } | ConvertTo-Json -Compress
      exit 0
    }
    throw
  }

  $definition = $task.Definition
  if ($definition.Actions.Count -ne 1) {
    @{ status = 'inaccessible' } | ConvertTo-Json -Compress
    exit 0
  }
  $taskAction = $definition.Actions.Item(1)
  $dailyAtLocal = $null
  if ($definition.Triggers.Count -eq 1 -and $definition.Triggers.Item(1).Type -eq 2) {
    $startBoundary = [string]$definition.Triggers.Item(1).StartBoundary
    if ($startBoundary -match 'T([0-9]{2}:[0-9]{2})') { $dailyAtLocal = $Matches[1] }
  }
  $multipleInstances = if ([int]$definition.Settings.MultipleInstances -eq 2) { 'IgnoreNew' } else { 'Other' }
  @{
    status = 'found'
    task = @{
      name = $TaskName
      execute = [string]$taskAction.Path
      workingDirectory = [string]$taskAction.WorkingDirectory
      arguments = [string]$taskAction.Arguments
      enabled = [bool]$task.Enabled
      running = ([int]$task.State -eq 4)
      taskState = [int]$task.State
      dailyAtLocal = $dailyAtLocal
      multipleInstances = $multipleInstances
      lastTaskResult = ([int64]$task.LastTaskResult -band [int64]4294967295)
    }
  } | ConvertTo-Json -Compress -Depth 5
  exit 0
}

if ($TaskName -ne 'StockReporterAssistant-ConferenceDaily') {
  throw 'Only the v2-owned task may be changed.'
}
if ($Action -eq 'MigrateTime') {
  if (-not [IO.Path]::IsPathRooted($Executable) -or [IO.Path]::GetFileName($Executable) -ne 'Update.exe' -or
      -not [IO.Path]::IsPathRooted($BackupDirectory)) {
    throw 'Invalid v2 task owner or backup directory.'
  }
  $task = $folder.GetTask($TaskName)
  $definition = $task.Definition
  if (-not [bool]$task.Enabled -or [int]$task.State -eq 4 -or
      $definition.Actions.Count -ne 1 -or $definition.Triggers.Count -ne 1 -or
      $definition.Triggers.Item(1).Type -ne 2 -or
      [int]$definition.Triggers.Item(1).DaysInterval -ne 1 -or
      [int]$definition.Settings.MultipleInstances -ne 2 -or
      [int]$definition.Principal.LogonType -ne 3 -or
      -not (Test-CurrentUserPrincipal ([string]$definition.Principal.UserId))) {
    throw 'The existing v2 task identity or daily trigger is not verified.'
  }
  $taskAction = $definition.Actions.Item(1)
  $taskPath = [IO.Path]::GetFullPath([string]$taskAction.Path)
  $ownerPath = [IO.Path]::GetFullPath($Executable)
  $taskDirectory = [IO.Path]::GetFullPath([string]$taskAction.WorkingDirectory)
  if (-not [string]::Equals($taskPath, $ownerPath, [StringComparison]::OrdinalIgnoreCase) -or
      -not [string]::Equals($taskDirectory, [IO.Path]::GetDirectoryName($ownerPath), [StringComparison]::OrdinalIgnoreCase) -or
      [string]$taskAction.Arguments -ne '--processStart StockReporterAssistant.exe --process-start-args "--conference-scheduled"' -or
      [string]$definition.Triggers.Item(1).StartBoundary -notmatch 'T17:45:00([+-][0-9]{2}:[0-9]{2})?$') {
    throw 'The existing v2 task is not the expected 17:45 task.'
  }
  $AllowDisabledLegacy = $true
  Assert-LegacyTaskAbsent $folder
  $beforeXml = [string]$task.Xml
  New-Item -ItemType Directory -Path $BackupDirectory -Force | Out-Null
  $backupPath = Join-Path $BackupDirectory ('conference-task-before-1900-' + [guid]::NewGuid().ToString('N') + '.xml')
  [IO.File]::WriteAllText($backupPath, $beforeXml, [Text.UTF8Encoding]::new($false))
  if ([IO.File]::ReadAllText($backupPath) -cne $beforeXml) { throw 'The v2 task backup could not be verified.' }
  $userId = [string]$definition.Principal.UserId
  $definition.Triggers.Item(1).StartBoundary = [string]$definition.Triggers.Item(1).StartBoundary -replace 'T17:45:00(?=([+-][0-9]{2}:[0-9]{2})?$)', 'T19:00:00'
  try {
    $folder.RegisterTaskDefinition($TaskName, $definition, 6, $userId, $null, 3) | Out-Null
    $verified = $folder.GetTask($TaskName)
    if ([string]$verified.Definition.Triggers.Item(1).StartBoundary -notmatch 'T19:00:00([+-][0-9]{2}:[0-9]{2})?$' -or
        -not [bool]$verified.Enabled) { throw 'The v2 task did not retain its daily 19:00 trigger.' }
    Assert-LegacyTaskAbsent $folder
  } catch {
    $failure = $_
    $folder.RegisterTask($TaskName, $beforeXml, 6, $userId, $null, 3) | Out-Null
    throw $failure
  }
  @{ status = 'migrated'; backup = $backupPath } | ConvertTo-Json -Compress
  exit 0
}
if ($Action -eq 'Remove') {
  if (-not [IO.Path]::IsPathRooted($Executable) -or [IO.Path]::GetFileName($Executable) -ne 'Update.exe') {
    throw 'Invalid v2 task owner path.'
  }
  try {
    $task = $folder.GetTask($TaskName)
  } catch {
    if (Test-MissingTaskException $_.Exception) {
      @{ status = 'skipped' } | ConvertTo-Json -Compress
      exit 0
    }
    throw
  }
  $definition = $task.Definition
  if ($definition.Actions.Count -ne 1 -or $definition.Triggers.Count -ne 1 -or
      $definition.Triggers.Item(1).Type -ne 2 -or
      [int]$definition.Settings.MultipleInstances -ne 2 -or
      [int]$definition.Principal.LogonType -ne 3) {
    @{ status = 'skipped' } | ConvertTo-Json -Compress
    exit 0
  }
  $taskAction = $definition.Actions.Item(1)
  $taskPath = [IO.Path]::GetFullPath([string]$taskAction.Path)
  $ownerPath = [IO.Path]::GetFullPath($Executable)
  $expectedDirectory = [IO.Path]::GetDirectoryName($ownerPath)
  $taskDirectory = [IO.Path]::GetFullPath([string]$taskAction.WorkingDirectory)
  $startBoundary = [string]$definition.Triggers.Item(1).StartBoundary
  if (-not [string]::Equals($taskPath, $ownerPath, [StringComparison]::OrdinalIgnoreCase) -or
      -not [string]::Equals($taskDirectory, $expectedDirectory, [StringComparison]::OrdinalIgnoreCase) -or
      $taskAction.Arguments -ne '--processStart StockReporterAssistant.exe --process-start-args "--conference-scheduled"' -or
      -not (Test-CurrentUserPrincipal ([string]$definition.Principal.UserId)) -or
      $startBoundary -notmatch '(T17:45:00|T19:00:00)([+-][0-9]{2}:[0-9]{2})?$') {
    @{ status = 'skipped' } | ConvertTo-Json -Compress
    exit 0
  }
  $folder.DeleteTask($TaskName, 0)
  @{ status = 'removed' } | ConvertTo-Json -Compress
  exit 0
}
if (-not [IO.Path]::IsPathRooted($Executable) -or
    [IO.Path]::GetFileName($Executable) -ne 'Update.exe' -or
    [IO.Path]::GetFullPath($WorkingDirectory) -ne [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Executable))) {
  throw 'Invalid upgrade-stable task executable or working directory.'
}

$definition = $service.NewTask(0)
$definition.RegistrationInfo.Description = 'Daily stock reporter conference Calendar check'
$definition.Settings.Enabled = $true
$definition.Settings.MultipleInstances = 2
$definition.Settings.StartWhenAvailable = $true
$definition.Settings.DisallowStartIfOnBatteries = $false
$definition.Settings.StopIfGoingOnBatteries = $false
$definition.Principal.LogonType = 3
$userId = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$definition.Principal.UserId = $userId

$trigger = $definition.Triggers.Create(2)
$trigger.StartBoundary = (Get-Date -Format 'yyyy-MM-dd') + 'T19:00:00'
$trigger.DaysInterval = 1
$taskAction = $definition.Actions.Create(0)
$taskAction.Path = $Executable
$taskAction.WorkingDirectory = $WorkingDirectory
$taskAction.Arguments = $Arguments

# Recheck the legacy writer at the last possible point. Any unknown state fails closed.
Assert-LegacyTaskAbsent $folder

# TASK_CREATE (2) fails if a task with the same name appears after inspection.
$folder.RegisterTaskDefinition($TaskName, $definition, 2, $userId, $null, 3) | Out-Null
@{ status = 'registered' } | ConvertTo-Json -Compress
