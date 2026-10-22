<#
.SYNOPSIS
    Schedules the monthly report in Windows Task Scheduler.

.PARAMETER ConfigFile
    Path to the client configuration file (stored as an absolute path in the task).

.PARAMETER Day
    Day of the month to run. Default: 1 (the report covers the previous month).

.PARAMETER Hour
    Time to run. Default: 09:00

.PARAMETER SendEmail
    Send the report by email automatically.

.EXAMPLE
    .\Schedule-MonthlyReport.ps1 -ConfigFile ".\config\contoso.json" -Day 1 -Hour "09:00" -SendEmail
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigFile,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 31)]
    [int]$Day = 1,

    [Parameter(Mandatory = $false)]
    [string]$Hour = "09:00",

    [Parameter(Mandatory = $false)]
    [switch]$SendEmail
)

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$GetReportScript = Join-Path $ScriptPath "Get-M365Report.ps1"

if (!(Test-Path $ConfigFile)) {
    Write-Host "[ERROR] Config file not found: $ConfigFile" -ForegroundColor Red
    exit 1
}
if ($Day -gt 28) {
    Write-Host "[WARN] Day $Day does not exist in every month; those months will be skipped. Day 1 is recommended." -ForegroundColor Yellow
}

# The task runs with a different working directory, so store an absolute path
$ConfigFile = (Resolve-Path $ConfigFile).Path

# Read the config to get the client name (English keys first, legacy keys as fallback)
$config = Get-Content $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
$displayName = if ($config.client.name) { $config.client.name } else { $config.cliente.nome }
$clientName = $displayName -replace '[^a-zA-Z0-9]', '_'

$taskName = "M365_Monthly_Report_$clientName"

# Prefer PowerShell 7 (pwsh.exe) when installed, otherwise Windows PowerShell
$pwsh = Get-Command "pwsh.exe" -ErrorAction SilentlyContinue
$executable = if ($pwsh) { $pwsh.Source } else { "powershell.exe" }

$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$GetReportScript`" -ConfigFile `"$ConfigFile`""
if ($SendEmail) {
    $arguments += " -SendEmail"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " SCHEDULING TASK IN TASK SCHEDULER" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Client:     $displayName"
Write-Host "Task:       $taskName"
Write-Host "Day:        $Day of every month"
Write-Host "Time:       $Hour"
Write-Host "Send email: $SendEmail"
Write-Host "Runs with:  $executable"
Write-Host "Config:     $ConfigFile"
Write-Host ""

# New-ScheduledTaskTrigger has no monthly option. Register with a daily trigger,
# then swap its schedule for a monthly one in the task XML (see the try block below).
$trigger = New-ScheduledTaskTrigger -Daily -At $Hour
$action = New-ScheduledTaskAction -Execute $executable -Argument $arguments -WorkingDirectory $ScriptPath

# Run as the current user, whether logged on or not
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U -RunLevel Limited

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Task already exists. Updating..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

try {
    Register-ScheduledTask -TaskName $taskName -Trigger $trigger -Action $action -Principal $principal -Settings $settings -Description "Monthly M365 report - $displayName" -ErrorAction Stop | Out-Null

    $months  = '<January/><February/><March/><April/><May/><June/><July/><August/><September/><October/><November/><December/>'
    $monthly = "<ScheduleByMonth><DaysOfMonth><Day>$Day</Day></DaysOfMonth><Months>$months</Months></ScheduleByMonth>"
    $xml = Export-ScheduledTask -TaskName $taskName -ErrorAction Stop
    $newXml = [regex]::Replace($xml, '<ScheduleByDay>.*?</ScheduleByDay>', $monthly, 'Singleline')
    if ($newXml -eq $xml) { throw "Could not convert the daily trigger to monthly (ScheduleByDay not found in the task XML)." }
    Register-ScheduledTask -TaskName $taskName -Xml $newXml -Force -ErrorAction Stop | Out-Null

    Write-Host ""
    Write-Host "[OK] Task scheduled successfully." -ForegroundColor Green
    Write-Host ""
    Write-Host "Check:      Get-ScheduledTask -TaskName '$taskName'" -ForegroundColor Cyan
    Write-Host "Run now:    Start-ScheduledTask -TaskName '$taskName'" -ForegroundColor Cyan
    Write-Host "Remove:     Unregister-ScheduledTask -TaskName '$taskName'" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Note: the task uses S4U logon, which may not be able to read user-profile protected keys" -ForegroundColor Yellow
    Write-Host "      (e.g. CurrentUser certificate private keys). If certificate auth fails in the task," -ForegroundColor Yellow
    Write-Host "      use a certificate in LocalMachine\My that this account can read." -ForegroundColor Yellow
    Write-Host ""
}
catch {
    Write-Host "[ERROR] Failed to create the task: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Try running as Administrator." -ForegroundColor Yellow
}
