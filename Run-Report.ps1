<#
.SYNOPSIS
    Quick runner - runs the report for one client or for all clients.

.PARAMETER Client
    Config file name without extension (e.g. "contoso" for config\contoso.json).
    If omitted (and -All is not used), lists the configured clients.

.PARAMETER All
    Runs the report for EVERY client in the config folder (sample.json is skipped).

.PARAMETER Preview
    Opens the HTML in the browser instead of generating a PDF.

.PARAMETER SendEmail
    Sends the email after generating the PDF.

.EXAMPLE
    .\Run-Report.ps1 -Client contoso -Preview

.EXAMPLE
    .\Run-Report.ps1 -Client contoso -SendEmail

.EXAMPLE
    .\Run-Report.ps1 -All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Client,

    [Parameter(Mandatory = $false)]
    [switch]$All,

    [Parameter(Mandatory = $false)]
    [switch]$Preview,

    [Parameter(Mandatory = $false)]
    [switch]$SendEmail
)

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptPath "config"
$GetReportScript = Join-Path $ScriptPath "Get-M365Report.ps1"

# Configured clients (the sample file is not a real client)
$configs = @(Get-ChildItem -Path $ConfigPath -Filter "*.json" -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -ne "sample.json" })

function Get-ClientDisplayName {
    param([string]$Path)
    try {
        $data = Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($data.client.name) { return $data.client.name }
        return $data.cliente.nome
    }
    catch { return "(invalid JSON)" }
}

if (!$Client -and !$All) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " M365 MONTHLY REPORT - CLIENTS" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    if ($configs.Count -eq 0) {
        Write-Host "No clients configured." -ForegroundColor Yellow
        Write-Host "Use New-ClientConfig.ps1 to add one." -ForegroundColor Yellow
    }
    else {
        Write-Host "Available clients:" -ForegroundColor Green
        Write-Host ""
        foreach ($cfg in $configs) {
            Write-Host "  - $($cfg.BaseName)" -ForegroundColor White -NoNewline
            Write-Host " ($(Get-ClientDisplayName $cfg.FullName))" -ForegroundColor Gray
        }
        Write-Host ""
        Write-Host "Usage:" -ForegroundColor Yellow
        Write-Host "  .\Run-Report.ps1 -Client <name> -Preview    # Open in the browser"
        Write-Host "  .\Run-Report.ps1 -Client <name>             # Generate PDF"
        Write-Host "  .\Run-Report.ps1 -Client <name> -SendEmail  # Generate and send"
        Write-Host "  .\Run-Report.ps1 -All                       # All clients"
    }
    Write-Host ""
    exit 0
}

# Single client
if ($Client) {
    $configFile = Join-Path $ConfigPath "$Client.json"

    if (!(Test-Path $configFile)) {
        Write-Host "Client not found: $Client" -ForegroundColor Red
        Write-Host "Expected file: $configFile" -ForegroundColor Yellow
        exit 1
    }

    Write-Host ""
    Write-Host "Running report for: $Client" -ForegroundColor Cyan
    Write-Host ""

    $params = @{
        ConfigFile = $configFile
    }

    if ($Preview) { $params.Preview = $true }
    if ($SendEmail) { $params.SendEmail = $true }

    & $GetReportScript @params
}

# All clients
if ($All) {
    if ($configs.Count -eq 0) {
        Write-Host "No clients configured." -ForegroundColor Yellow
        exit 1
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " RUNNING FOR ALL CLIENTS" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    $results = @()

    foreach ($cfg in $configs) {
        Write-Host "----------------------------------------" -ForegroundColor Gray
        Write-Host "Client: $($cfg.BaseName)" -ForegroundColor Cyan

        try {
            $params = @{
                ConfigFile = $cfg.FullName
            }

            if ($SendEmail) { $params.SendEmail = $true }

            $pdfPath = & $GetReportScript @params

            $results += [PSCustomObject]@{
                Client = $cfg.BaseName
                Status = "OK"
                PDF    = $pdfPath
            }
        }
        catch {
            $results += [PSCustomObject]@{
                Client = $cfg.BaseName
                Status = "ERROR"
                PDF    = $_.Exception.Message
            }
            Write-Host "ERROR: $_" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host " RUN SUMMARY" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""

    $results | Format-Table -AutoSize
}
