Function Check-RunAsAdministrator()
{
  #Get current user context
  $CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
  
  #Check user is running the script is member of Administrator Group
  if($CurrentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
  {
       Write-host "Script is running with Administrator privileges!"
  }
  else
    {
       #Create a new Elevated process to Start PowerShell
       $ElevatedProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell";
 
       # Specify the current script path and name as a parameter
       $ElevatedProcess.Arguments = "& '" + $script:MyInvocation.MyCommand.Path + "'"
 
       #Set the Process to elevated
       $ElevatedProcess.Verb = "runas"
 
       #Start the new elevated process
       [System.Diagnostics.Process]::Start($ElevatedProcess)
 
       #Exit from the current, unelevated, process
       Exit
 
    }
}
 
#Check Script is running with Elevated Privileges
Check-RunAsAdministrator
 
#Place your script here.

$currentDir = $PSScriptRoot
$exePath = $currentDir+"\stock-info-collector.exe"

$trigger = New-ScheduledTaskTrigger -Daily -At 7pm
$action = New-ScheduledTaskAction -Execute $exePath -WorkingDirectory $PSScriptRoot

# 新增設定執行層級為最高，不論使用者是否登入
$principal = New-ScheduledTaskPrincipal -UserId "User" -LogonType Password -RunLevel Highest

# 設定工作失敗後重試行為
$settings = New-ScheduledTaskSettingsSet -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5)

# 請求使用者輸入密碼
#$password = Read-Host -Prompt "請輸入使用者的密碼" -AsSecureString

#Register-ScheduledTask -TaskName "stock-info-collector-daily-work" -Trigger $trigger -Action $action -Description '每日法說爬蟲並新增Google行事曆。' -Principal $principal -Settings $settings -User 'User' -Password $password
Register-ScheduledTask -TaskName "stock-info-collector-daily-work" -Trigger $trigger -Action $action -Description '每日法說爬蟲並新增Google行事曆。'  -Settings $settings #-Principal $principal




pause