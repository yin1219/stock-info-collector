$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'installer-preflight-core.ps1')

$setup = Join-Path $PSScriptRoot 'StockReporterAssistantSquirrelSetup.exe'
if (-not (Test-Path -LiteralPath $setup -PathType Leaf)) {
  [System.Windows.Forms.MessageBox]::Show('安裝內容不完整，無法繼續。', '股市記者小幫手安裝', 'OK', 'Error') | Out-Null
  exit 1
}

try {
  $running = @(Get-CimInstance Win32_Process -Filter "Name = 'StockReporterAssistant.exe'" -ErrorAction Stop)
  $outcome = Invoke-ReporterInstallerPlan $running $env:LOCALAPPDATA {
    param($count)
    $message = "偵測到 $count 個已安裝的股市記者小幫手正在執行（可能縮在系統匣）。`n`n要先結束舊版並繼續安裝嗎？進行中的檢查或同步會中斷；選擇「否」會取消安裝。"
    [System.Windows.Forms.MessageBox]::Show($message, '股市記者小幫手更新', 'YesNo', 'Warning', 'Button2') -eq 'Yes'
  } {
    param($processId)
    Stop-Process -Id $processId -Force -ErrorAction Stop
  } {
    param($processId)
    try { Wait-Process -Id $processId -Timeout 15 -ErrorAction Stop; $true }
    catch {
      # A process that has already exited is also safe to proceed.
      -not (Get-Process -Id $processId -ErrorAction SilentlyContinue)
    }
  } {
    $installer = Start-Process -FilePath $setup -PassThru -ErrorAction Stop
    $installer.WaitForExit()
    if ($installer.ExitCode -ne 0) { throw "安裝程式結束碼：$($installer.ExitCode)" }
  }
  if ($outcome -eq 'unknown') {
    [System.Windows.Forms.MessageBox]::Show('無法確認執行中的小幫手位置。請先從系統匣手動結束，再重新安裝。', '股市記者小幫手更新', 'OK', 'Warning') | Out-Null
    exit 2
  }
  if ($outcome -eq 'close_failed') { throw '舊版未能完全結束；未啟動安裝。' }
  if ($outcome -eq 'cancelled') { exit 0 }
} catch {
  [System.Windows.Forms.MessageBox]::Show("安裝未完成：$($_.Exception.Message)", '股市記者小幫手更新', 'OK', 'Error') | Out-Null
  exit 1
}
