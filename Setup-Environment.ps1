<#
.SYNOPSIS
    Configura o ambiente para executar os relatorios M365.

.DESCRIPTION
    - Verifica pre-requisitos
    - Instala modulos necessarios
    - Testa conectividade
#>

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " M365 MONTHLY REPORT - SETUP" -ForegroundColor Cyan
Write-Host " Casco Digital MSP Tools" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Verificar PowerShell version
$psVersion = $PSVersionTable.PSVersion
Write-Host "[CHECK] PowerShell Version: $($psVersion.Major).$($psVersion.Minor)" -ForegroundColor Green

# Verificar Edge
$edgePaths = @(
    "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
)
$edgeFound = $false
foreach ($path in $edgePaths) {
    if (Test-Path $path) {
        Write-Host "[CHECK] Microsoft Edge encontrado: $path" -ForegroundColor Green
        $edgeFound = $true
        break
    }
}
if (!$edgeFound) {
    Write-Host "[WARN] Microsoft Edge nao encontrado. Necessario para gerar PDFs." -ForegroundColor Yellow
}

# Verificar/criar estrutura de pastas
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$folders = @("config", "output", "logs", "templates")
foreach ($folder in $folders) {
    $folderPath = Join-Path $ScriptPath $folder
    if (!(Test-Path $folderPath)) {
        New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
        Write-Host "[CREATED] Pasta criada: $folder" -ForegroundColor Green
    }
    else {
        Write-Host "[OK] Pasta existe: $folder" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host " CONFIGURACAO DO APP REGISTRATION" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "Para cada tenant (seu e de clientes), crie um App Registration:"
Write-Host ""
Write-Host "1. Acesse: https://entra.microsoft.com" -ForegroundColor Cyan
Write-Host "2. Va em: Identity > Applications > App registrations"
Write-Host "3. Clique: + New registration"
Write-Host "4. Nome: 'MSP Monthly Report' (ou similar)"
Write-Host "5. Supported account types: 'Single tenant'"
Write-Host "6. Redirect URI: deixe vazio"
Write-Host "7. Clique: Register"
Write-Host ""
Write-Host "Apos criar, configure as PERMISSOES (API permissions):" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Microsoft Graph (Application permissions):" -ForegroundColor Cyan
Write-Host "  - AuditLog.Read.All          (sign-in logs)"
Write-Host "  - Directory.Read.All         (usuarios, roles)"
Write-Host "  - Reports.Read.All           (usage reports)"
Write-Host "  - SecurityEvents.Read.All    (secure score)"
Write-Host "  - User.Read.All              (lista usuarios)"
Write-Host "  - Mail.Send                  (enviar email - so no SEU tenant)"
Write-Host ""
Write-Host "IMPORTANTE: Clique em 'Grant admin consent' apos adicionar!" -ForegroundColor Red
Write-Host ""
Write-Host "Depois, crie um CLIENT SECRET:" -ForegroundColor Yellow
Write-Host "  1. Va em: Certificates & secrets"
Write-Host "  2. + New client secret"
Write-Host "  3. Descricao: 'MSP Report'"
Write-Host "  4. Expiracao: 24 meses"
Write-Host "  5. COPIE O VALUE (so aparece uma vez!)"
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host " SHARED MAILBOX (seu tenant)" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "Para enviar emails de suporte@seudominio.com:"
Write-Host ""
Write-Host "1. Exchange Admin Center > Recipients > Shared"
Write-Host "2. Crie ou use shared mailbox existente"
Write-Host "3. Adicione seu usuario como membro"
Write-Host "4. O App Registration do SEU tenant precisa de Mail.Send"
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " PROXIMO PASSO" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Apos criar o App Registration, execute:" -ForegroundColor Cyan
Write-Host ""
Write-Host '.\New-ClientConfig.ps1 -ClientName "Nome Cliente" -TenantId "xxx" -ClientId "yyy" -ClientSecret "zzz" -Email "email@cliente.com"'
Write-Host ""
Write-Host "Ou edite manualmente: config\cascodigital.json"
Write-Host ""
