<#
.SYNOPSIS
    Script de execucao rapida - roda relatorio para um ou todos os clientes.

.PARAMETER Client
    Nome do arquivo de config (sem extensao). Ex: "cascodigital"
    Se nao especificado, lista clientes disponiveis.

.PARAMETER All
    Se especificado, roda para TODOS os clientes na pasta config.

.PARAMETER Preview
    Abre HTML no navegador ao inves de gerar PDF.

.PARAMETER SendEmail
    Envia email apos gerar PDF.

.EXAMPLE
    .\Run-Report.ps1 -Client cascodigital -Preview

.EXAMPLE
    .\Run-Report.ps1 -Client cascodigital -SendEmail

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

# Listar clientes disponiveis
$configs = Get-ChildItem -Path $ConfigPath -Filter "*.json" -ErrorAction SilentlyContinue

if (!$Client -and !$All) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " M365 MONTHLY REPORT - CLIENTES" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    if ($configs.Count -eq 0) {
        Write-Host "Nenhum cliente configurado." -ForegroundColor Yellow
        Write-Host "Use New-ClientConfig.ps1 para adicionar." -ForegroundColor Yellow
    }
    else {
        Write-Host "Clientes disponiveis:" -ForegroundColor Green
        Write-Host ""
        foreach ($cfg in $configs) {
            $data = Get-Content $cfg.FullName -Raw | ConvertFrom-Json
            Write-Host "  - $($cfg.BaseName)" -ForegroundColor White -NoNewline
            Write-Host " ($($data.cliente.nome))" -ForegroundColor Gray
        }
        Write-Host ""
        Write-Host "Uso:" -ForegroundColor Yellow
        Write-Host "  .\Run-Report.ps1 -Client <nome> -Preview    # Abre no navegador"
        Write-Host "  .\Run-Report.ps1 -Client <nome>             # Gera PDF"
        Write-Host "  .\Run-Report.ps1 -Client <nome> -SendEmail  # Gera e envia"
        Write-Host "  .\Run-Report.ps1 -All                       # Todos os clientes"
    }
    Write-Host ""
    exit 0
}

# Executar para um cliente especifico
if ($Client) {
    $configFile = Join-Path $ConfigPath "$Client.json"

    if (!(Test-Path $configFile)) {
        Write-Host "Cliente nao encontrado: $Client" -ForegroundColor Red
        Write-Host "Arquivo esperado: $configFile" -ForegroundColor Yellow
        exit 1
    }

    Write-Host ""
    Write-Host "Executando relatorio para: $Client" -ForegroundColor Cyan
    Write-Host ""

    $params = @{
        ConfigFile = $configFile
    }

    if ($Preview) { $params.Preview = $true }
    if ($SendEmail) { $params.SendEmail = $true }

    & $GetReportScript @params
}

# Executar para todos os clientes
if ($All) {
    if ($configs.Count -eq 0) {
        Write-Host "Nenhum cliente configurado." -ForegroundColor Yellow
        exit 1
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " EXECUTANDO PARA TODOS OS CLIENTES" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    $results = @()

    foreach ($cfg in $configs) {
        Write-Host "----------------------------------------" -ForegroundColor Gray
        Write-Host "Cliente: $($cfg.BaseName)" -ForegroundColor Cyan

        try {
            $params = @{
                ConfigFile = $cfg.FullName
            }

            if ($SendEmail) { $params.SendEmail = $true }

            $pdfPath = & $GetReportScript @params

            $results += [PSCustomObject]@{
                Cliente = $cfg.BaseName
                Status  = "OK"
                PDF     = $pdfPath
            }
        }
        catch {
            $results += [PSCustomObject]@{
                Cliente = $cfg.BaseName
                Status  = "ERRO"
                PDF     = $_.Exception.Message
            }
            Write-Host "ERRO: $_" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host " RESUMO DA EXECUCAO" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""

    $results | Format-Table -AutoSize
}
