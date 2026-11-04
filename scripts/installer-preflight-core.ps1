function Test-ReporterInstalledExecutable {
  param([string]$ExecutablePath, [string]$LocalAppData)

  if ([string]::IsNullOrWhiteSpace($ExecutablePath) -or [string]::IsNullOrWhiteSpace($LocalAppData)) { return $false }
  $root = [regex]::Escape($LocalAppData.TrimEnd('\'))
  $leaf = 'StockReporterAssistant\\app-\d+\.\d+\.\d+\\StockReporterAssistant\.exe$'
  return $ExecutablePath -match "^$root\\$leaf" -or
    $ExecutablePath -match "^$root\\Packages\\[^\\]+\\LocalCache\\Local\\$leaf"
}

function Invoke-ReporterInstallerPlan {
  param(
    [object[]]$Processes,
    [string]$LocalAppData,
    [scriptblock]$ConfirmClose,
    [scriptblock]$StopProcess,
    [scriptblock]$WaitForExit,
    [scriptblock]$LaunchSetup
  )

  $verified = @()
  foreach ($candidate in $Processes) {
    if ([string]::IsNullOrWhiteSpace($candidate.ExecutablePath)) { return 'unknown' }
    if (Test-ReporterInstalledExecutable $candidate.ExecutablePath $LocalAppData) {
      $verified += $candidate
    }
  }
  if ($verified.Count -gt 0) {
    if (-not (& $ConfirmClose $verified.Count)) { return 'cancelled' }
    foreach ($candidate in $verified) {
      try { & $StopProcess $candidate.ProcessId }
      catch {
        if (-not (& $WaitForExit $candidate.ProcessId)) { return 'close_failed' }
      }
    }
    foreach ($candidate in $verified) {
      if (-not (& $WaitForExit $candidate.ProcessId)) { return 'close_failed' }
    }
  }
  & $LaunchSetup | Out-Null
  return 'installed'
}
