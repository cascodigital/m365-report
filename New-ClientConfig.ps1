<#
.SYNOPSIS
    Cria arquivo de configuracao para um novo cliente.

.PARAMETER ClientName
    Nome do cliente (usado no nome do arquivo e no relatorio).

.PARAMETER TenantId
    Tenant ID do cliente (Azure AD).

.PARAMETER ClientId
    Client ID do App Registration.

.PARAMETER ClientSecret
    Client Secret do App Registration.

.PARAMETER Email
    Email do destinatario do relatorio.

.EXAMPLE
    .\New-ClientConfig.ps1 -ClientName "Empresa XYZ" -TenantId "xxx-xxx" -ClientId "yyy-yyy" -ClientSecret "zzz" -Email "ti@empresa.com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientName,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $true)]
    [string]$Email,

    [Parameter(Mandatory = $false)]
    [string]$MSPName = "Casco Digital",

    [Parameter(Mandatory = $false)]
    [string]$MSPContact = "suporte@cascodigital.com.br",

    [Parameter(Mandatory = $false)]
    [string]$MSPSender = "suporte@cascodigital.com.br"
)

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptPath "config"

if (!(Test-Path $ConfigPath)) {
    New-Item -ItemType Directory -Path $ConfigPath -Force | Out-Null
}

$safeFileName = $ClientName -replace '[^a-zA-Z0-9]', '_'
$configFile = Join-Path $ConfigPath "$safeFileName.json"

$config = @{
    cliente = @{
        nome          = $ClientName
        tenant_id     = $TenantId
        client_id     = $ClientId
        client_secret = $ClientSecret
    }
    email = @{
        destinatarios     = @($Email)
        cc                = @()
        assunto_template  = "Relatorio Mensal Microsoft 365 - {{PERIODO}}"
    }
    msp = @{
        nome      = $MSPName
        contato   = $MSPContact
        remetente = $MSPSender
    }
    opcoes = @{
        incluir_usuarios_sem_mfa = $true
        incluir_recomendacoes    = $true
        idioma                   = "pt-BR"
    }
}

$config | ConvertTo-Json -Depth 5 | Out-File -FilePath $configFile -Encoding UTF8

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " CONFIGURACAO CRIADA COM SUCESSO" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Arquivo: $configFile" -ForegroundColor Cyan
Write-Host ""
Write-Host "Para testar (preview sem enviar email):" -ForegroundColor Yellow
Write-Host "  .\Get-M365Report.ps1 -ConfigFile `"$configFile`" -Preview" -ForegroundColor White
Write-Host ""
Write-Host "Para gerar PDF:" -ForegroundColor Yellow
Write-Host "  .\Get-M365Report.ps1 -ConfigFile `"$configFile`"" -ForegroundColor White
Write-Host ""
Write-Host "Para gerar e enviar por email:" -ForegroundColor Yellow
Write-Host "  .\Get-M365Report.ps1 -ConfigFile `"$configFile`" -SendEmail" -ForegroundColor White
Write-Host ""

return $configFile
