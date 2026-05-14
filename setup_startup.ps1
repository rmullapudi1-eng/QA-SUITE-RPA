# Run this once as Administrator to register QA Suite as a Windows startup task
# Right-click setup_startup.ps1 → Run with PowerShell (as Administrator)

$taskName   = "QA Suite RPA"
$appDir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$python     = "C:\Program Files\Python311\python.exe"
$script     = Join-Path $appDir "app_ui.py"

# Remove old task if exists
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

# Trigger: at user logon
$trigger = New-ScheduledTaskTrigger -AtLogOn

# Action: run python app_ui.py in the app folder
$action = New-ScheduledTaskAction `
    -Execute $python `
    -Argument "`"$script`"" `
    -WorkingDirectory $appDir

# Settings: run only when logged on, restart on failure
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 0) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $taskName `
    -Trigger $trigger `
    -Action $action `
    -Settings $settings `
    -RunLevel Highest `
    -Force | Out-Null

Write-Host ""
Write-Host "SUCCESS: '$taskName' registered as a startup task." -ForegroundColor Green
Write-Host "The app will auto-start every time you log into Windows." -ForegroundColor Green
Write-Host ""
Write-Host "Scheduled RPA runs inside the app:"
Write-Host "  - Mon to Fri : 7:00 PM IST"
Write-Host "  - Tue to Sat : 2:00 AM IST"
Write-Host ""
pause
