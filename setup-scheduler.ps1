# HOAi Daily Report — Windows Task Scheduler Setup
#
# Creates a scheduled task that runs at 7:00 AM ET daily:
#   1. fetch-daily-data.py (yesterday's data)
#   2. generate-daily-report.py (Excel + PDF)
#   3. Optional email send with both attachments
#
# Run once: powershell -ExecutionPolicy Bypass -File daily-reports\setup-scheduler.ps1

$ErrorActionPreference = "Stop"

$TaskName = "HOAi_Daily_Report"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonExe = "python"  # Assumes python is on PATH
$LogDir = Join-Path $ScriptDir "logs"

# Create log directory
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# Build the action script that runs both steps
$RunScript = @"
@echo off
cd /d "$ScriptDir\.."
echo [%date% %time%] Starting daily report fetch >> "$LogDir\daily-report.log"
$PythonExe "$ScriptDir\fetch-daily-data.py" >> "$LogDir\daily-report.log" 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] FETCH FAILED >> "$LogDir\daily-report.log"
    exit /b 1
)
echo [%date% %time%] Starting report generation >> "$LogDir\daily-report.log"
$PythonExe "$ScriptDir\generate-daily-report.py" >> "$LogDir\daily-report.log" 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] GENERATE FAILED >> "$LogDir\daily-report.log"
    exit /b 1
)
echo [%date% %time%] Sending email >> "$LogDir\daily-report.log"
$PythonExe "$ScriptDir\send-daily-email.py" >> "$LogDir\daily-report.log" 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] EMAIL SEND WARNING - continuing >> "$LogDir\daily-report.log"
)
echo [%date% %time%] Daily report complete >> "$LogDir\daily-report.log"
"@

$BatchPath = Join-Path $ScriptDir "run-daily-report.bat"
Set-Content -Path $BatchPath -Value $RunScript -Encoding ASCII

# Remove existing task if present
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed existing task: $TaskName"
}

# Create trigger: 7:00 AM daily
$Trigger = New-ScheduledTaskTrigger -Daily -At "07:00"

# Create action
$Action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument "/c `"$BatchPath`"" `
    -WorkingDirectory (Split-Path -Parent $ScriptDir)

# Settings
$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

# Register
Register-ScheduledTask `
    -TaskName $TaskName `
    -Trigger $Trigger `
    -Action $Action `
    -Settings $Settings `
    -Description "HOAi Daily Report: fetch data + generate Excel/PDF at 7 AM ET" `
    -RunLevel Highest

Write-Host ""
Write-Host "Scheduled task created: $TaskName"
Write-Host "  Trigger: Daily at 7:00 AM"
Write-Host "  Action:  $BatchPath"
Write-Host "  Logs:    $LogDir\daily-report.log"
Write-Host ""
Write-Host "To test manually:"
Write-Host "  cmd /c `"$BatchPath`""
Write-Host ""
Write-Host "To remove:"
Write-Host "  Unregister-ScheduledTask -TaskName $TaskName"
