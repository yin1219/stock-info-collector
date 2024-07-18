$currentDir = $PSScriptRoot
$exePath = $currentDir+"\stock-info-collector.exe"

$trigger = New-ScheduledTaskTrigger -Daily -At 7pm
$action = New-ScheduledTaskAction -Execute $exePath -WorkingDirectory $PSScriptRoot

# 設定工作失敗後重試行為
$settings = New-ScheduledTaskSettingsSet -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5)

Register-ScheduledTask -TaskName "stock-info-collector-daily-work" -Trigger $trigger -Action $action -Description '每日法說爬蟲並新增Google行事曆。'  -Settings $settings #-Principal $principal


