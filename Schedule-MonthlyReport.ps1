<#
.SYNOPSIS
    Agenda a execucao mensal do relatorio no Task Scheduler.

.PARAMETER ConfigFile
    Caminho do arquivo de configuracao do cliente.

.PARAMETER Day
    Dia do mes para executar. Default: 30

.PARAMETER Hour
    Hora para executar. Default: 09:00

.PARAMETER SendEmail
    Se deve enviar email automaticamente.

.EXAMPLE
    .\Schedule-MonthlyReport.ps1 -ConfigFile ".\config\cascodigital.json" -Day 30 -Hour "09:00"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigFile,

    [Parameter(Mandatory = $false)]
    [int]$Day = 30,

    [Parameter(Mandatory = $false)]
    [string]$Hour = "09:00",

    [Parameter(Mandatory = $false)]
    [switch]$SendEmail
)

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$GetReportScript = Join-Path $ScriptPath "Get-M365Report.ps1"

# Carregar config para pegar nome do cliente
$config = Get-Content $ConfigFile -Raw | ConvertFrom-Json
$clientName = $config.cliente.nome -replace '[^a-zA-Z0-9]', '_'

$taskName = "M365_Monthly_Report_$clientName"

# Montar argumentos
$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$GetReportScript`" -ConfigFile `"$ConfigFile`""
if ($SendEmail) {
    $arguments += " -SendEmail"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " AGENDANDO TAREFA NO TASK SCHEDULER" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Cliente: $($config.cliente.nome)"
Write-Host "Tarefa: $taskName"
Write-Host "Dia: $Day de cada mes"
Write-Host "Hora: $Hour"
Write-Host "Enviar Email: $SendEmail"
Write-Host ""

# Criar trigger mensal
$trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth $Day -At $Hour

# Criar action
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $arguments -WorkingDirectory $ScriptPath

# Criar principal (rodar como usuario atual)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U -RunLevel Limited

# Settings
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

# Verificar se tarefa ja existe
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

if ($existingTask) {
    Write-Host "Tarefa ja existe. Atualizando..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# Criar tarefa
try {
    Register-ScheduledTask -TaskName $taskName -Trigger $trigger -Action $action -Principal $principal -Settings $settings -Description "Relatorio mensal M365 - $($config.cliente.nome)" | Out-Null

    Write-Host ""
    Write-Host "[OK] Tarefa agendada com sucesso!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Para verificar: Get-ScheduledTask -TaskName '$taskName'" -ForegroundColor Cyan
    Write-Host "Para executar manualmente: Start-ScheduledTask -TaskName '$taskName'" -ForegroundColor Cyan
    Write-Host "Para remover: Unregister-ScheduledTask -TaskName '$taskName'" -ForegroundColor Cyan
    Write-Host ""
}
catch {
    Write-Host "[ERRO] Falha ao criar tarefa: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Tente executar como Administrador." -ForegroundColor Yellow
}
