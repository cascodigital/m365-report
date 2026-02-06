<#
.SYNOPSIS
    Coleta dados do Microsoft 365 via Graph API e gera relatório PDF.

.DESCRIPTION
    Script para MSPs gerarem relatórios mensais automatizados de saúde do M365.
    Coleta: Secure Score, Logins, MFA, Licenças, Storage, Uso de Serviços.

.PARAMETER ConfigFile
    Caminho para o arquivo JSON de configuração do cliente.

.PARAMETER OutputPath
    Pasta onde o PDF será salvo. Default: ./output

.PARAMETER Preview
    Se especificado, abre o HTML no navegador ao invés de gerar PDF.

.PARAMETER SendEmail
    Se especificado, envia o relatório por email após gerar.

.EXAMPLE
    .\Get-M365Report.ps1 -ConfigFile ".\config\cascodigital.json" -Preview

.EXAMPLE
    .\Get-M365Report.ps1 -ConfigFile ".\config\cascodigital.json" -SendEmail
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigFile,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\output",

    [Parameter(Mandatory = $false)]
    [switch]$Preview,

    [Parameter(Mandatory = $false)]
    [switch]$SendEmail
)

# ============================================================================
# CONFIGURACAO
# ============================================================================

$ErrorActionPreference = "Stop"
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$TemplatePath = Join-Path $ScriptPath "templates\report.html"
$LogPath = Join-Path $ScriptPath "logs"

# Garante que as pastas existem
if (!(Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
if (!(Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }

# Log file
$LogFile = Join-Path $LogPath "report_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $logMessage

    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARN"  { Write-Host $logMessage -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }
}

# ============================================================================
# CARREGAR CONFIGURACAO
# ============================================================================

Write-Log "Iniciando coleta de dados M365..."
Write-Log "Arquivo de config: $ConfigFile"

if (!(Test-Path $ConfigFile)) {
    Write-Log "Arquivo de configuracao nao encontrado: $ConfigFile" "ERROR"
    exit 1
}

$Config = Get-Content $ConfigFile -Raw | ConvertFrom-Json
$ClienteNome = $Config.cliente.nome
$TenantId = $Config.cliente.tenant_id
$ClientId = $Config.cliente.client_id
$ClientSecret = $Config.cliente.client_secret

Write-Log "Cliente: $ClienteNome"
Write-Log "Tenant: $TenantId"

# ============================================================================
# AUTENTICACAO GRAPH API
# ============================================================================

function Get-GraphToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }

    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    try {
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        Write-Log "Falha ao obter token: $_" "ERROR"
        throw
    }
}

function Invoke-GraphRequest {
    param(
        [string]$Token,
        [string]$Endpoint,
        [string]$Method = "GET",
        [hashtable]$Body = $null
    )

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }

    $uri = "https://graph.microsoft.com/v1.0$Endpoint"

    try {
        if ($Body) {
            $response = Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers -Body ($Body | ConvertTo-Json -Depth 10)
        }
        else {
            $response = Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers
        }
        return $response
    }
    catch {
        Write-Log "Erro na requisicao Graph ($Endpoint): $_" "WARN"
        return $null
    }
}

function Invoke-GraphRequestBeta {
    param(
        [string]$Token,
        [string]$Endpoint,
        [string]$Method = "GET"
    )

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }

    $uri = "https://graph.microsoft.com/beta$Endpoint"

    try {
        $response = Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers
        return $response
    }
    catch {
        Write-Log "Erro na requisicao Graph Beta ($Endpoint): $_" "WARN"
        return $null
    }
}

# ============================================================================
# COLETA DE DADOS
# ============================================================================

Write-Log "Autenticando no Graph API..."
$Token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
Write-Log "Token obtido com sucesso" "SUCCESS"

# Periodo do relatorio
$dataAtual = Get-Date
$mesAnterior = $dataAtual.AddMonths(-1)
$periodo = $mesAnterior.ToString("MMMM yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR"))
$periodo = (Get-Culture).TextInfo.ToTitleCase($periodo)

Write-Log "Periodo do relatorio: $periodo"

# --- Secure Score ---
Write-Log "Coletando Secure Score..."
$secureScoreData = Invoke-GraphRequest -Token $Token -Endpoint "/security/secureScores?`$top=2"

$secureScore = 0
$secureScoreMax = 100
$secureScoreAtual = 0
$secureScoreVariacao = "N/A"

if ($secureScoreData -and $secureScoreData.value.Count -gt 0) {
    $latest = $secureScoreData.value[0]
    $secureScoreAtual = [math]::Round($latest.currentScore, 1)
    $secureScoreMax = [math]::Round($latest.maxScore, 1)
    $secureScore = [math]::Round(($latest.currentScore / $latest.maxScore) * 100, 0)

    if ($secureScoreData.value.Count -gt 1) {
        $previous = $secureScoreData.value[1]
        $previousPercent = [math]::Round(($previous.currentScore / $previous.maxScore) * 100, 0)
        $diff = $secureScore - $previousPercent
        if ($diff -ge 0) {
            $secureScoreVariacao = "+$diff%"
        }
        else {
            $secureScoreVariacao = "$diff%"
        }
    }
    Write-Log "Secure Score: $secureScore% ($secureScoreAtual/$secureScoreMax)" "SUCCESS"
}
else {
    Write-Log "Secure Score nao disponivel" "WARN"
}

# --- Usuarios ---
Write-Log "Coletando usuarios..."
$usersData = Invoke-GraphRequest -Token $Token -Endpoint "/users?`$filter=accountEnabled eq true&`$select=id,displayName,userPrincipalName,assignedLicenses&`$top=999"

$totalUsuarios = 0
$usuarios = @()

if ($usersData -and $usersData.value) {
    $usuarios = $usersData.value | Where-Object { $_.userPrincipalName -notlike "*#EXT#*" }
    $totalUsuarios = $usuarios.Count
    Write-Log "Total de usuarios ativos: $totalUsuarios" "SUCCESS"
}

# --- MFA Status ---
Write-Log "Coletando status MFA..."
$mfaUsers = @()
$usuariosSemMFA = @()
$usuariosComMFA = 0

# Para cada usuario, verificar MFA via authentication methods
foreach ($user in $usuarios) {
    try {
        $authMethods = Invoke-GraphRequest -Token $Token -Endpoint "/users/$($user.id)/authentication/methods"

        # MFA esta ativo se tiver mais que apenas password
        $hasMFA = $false
        if ($authMethods -and $authMethods.value) {
            $nonPasswordMethods = $authMethods.value | Where-Object { $_.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod' }
            $hasMFA = $nonPasswordMethods.Count -gt 0
        }

        if ($hasMFA) {
            $usuariosComMFA++
        }
        else {
            $usuariosSemMFA += $user.displayName
        }
    }
    catch {
        # Se falhar, assume sem MFA
        $usuariosSemMFA += $user.displayName
    }
}

$mfaPercent = if ($totalUsuarios -gt 0) { [math]::Round(($usuariosComMFA / $totalUsuarios) * 100, 0) } else { 0 }
Write-Log "MFA: $usuariosComMFA de $totalUsuarios ($mfaPercent%)" "SUCCESS"

# --- Sign-in Logs (ultimos 7 dias - limitacao do basic) ---
Write-Log "Coletando logs de login..."
$seteDiasAtras = (Get-Date).AddDays(-7).ToString("yyyy-MM-ddTHH:mm:ssZ")

$signInSuccess = "N/D"
$signInFailed = "N/D"
$signInSuccessEstimado = "N/D"
$signInFailedEstimado = "N/D"

$signInLogs = Invoke-GraphRequest -Token $Token -Endpoint "/auditLogs/signIns?`$filter=createdDateTime ge $seteDiasAtras&`$top=999"

if ($null -eq $signInLogs) {
    # 403 Forbidden - provavelmente requer Azure AD Premium P1/P2
    Write-Log "Sign-in logs nao disponiveis - requer Azure AD Premium P1/P2" "WARN"
}
elseif ($signInLogs.value) {
    $signInSuccess = ($signInLogs.value | Where-Object { $_.status.errorCode -eq 0 }).Count
    $signInFailed = ($signInLogs.value | Where-Object { $_.status.errorCode -ne 0 }).Count
    Write-Log "Logins (7 dias): $signInSuccess sucesso, $signInFailed falhos" "SUCCESS"

    # Estimar para 30 dias (aproximacao)
    $signInSuccessEstimado = [math]::Round($signInSuccess * 4.3, 0)
    $signInFailedEstimado = [math]::Round($signInFailed * 4.3, 0)
}
else {
    # Resposta vazia - nenhum login no periodo
    $signInSuccess = 0
    $signInFailed = 0
    $signInSuccessEstimado = 0
    $signInFailedEstimado = 0
    Write-Log "Logins (7 dias): Nenhum registro encontrado" "SUCCESS"
}

# --- Licencas ---
Write-Log "Coletando licencas..."
$licensesData = Invoke-GraphRequest -Token $Token -Endpoint "/subscribedSkus"

$licencasTotal = 0
$licencasUsadas = 0
$licencasDetalhes = @()

# SKUs gratuitos/internos para ignorar no relatorio
$skusIgnorar = @(
    "*FLOW_FREE*",           # Power Automate Free
    "*POWERAUTOMATE_FREE*",  # Power Automate Free (outro nome)
    "*TEAMS_EXPLORATORY*",   # Teams Exploratory
    "*POWERAPPS_VIRAL*",     # PowerApps Trial
    "*POWER_BI_STANDARD*",   # Power BI Free
    "*RIGHTSMANAGEMENT_ADHOC*", # RMS Ad-hoc
    "*WINDOWS_STORE*",       # Windows Store for Business
    "*STREAM*",              # Stream Trial
    "*FORMS_PRO*",           # Forms Pro Trial
    "*CCIBOTS*",             # Power Virtual Agents Trial
    "*CLIPCHAMP*",           # Clipchamp
    "*Dynamics_365*TRIAL*",  # Dynamics trials
    "*VISIOCLIENT_FIRSTLINE*" # Visio web free
)

if ($licensesData -and $licensesData.value) {
    foreach ($sku in $licensesData.value) {
        if ($sku.capabilityStatus -eq "Enabled") {

            # Verificar se deve ignorar este SKU
            $ignorar = $false
            foreach ($pattern in $skusIgnorar) {
                if ($sku.skuPartNumber -like $pattern) {
                    $ignorar = $true
                    Write-Log "Ignorando licenca gratuita: $($sku.skuPartNumber)" "INFO"
                    break
                }
            }

            if ($ignorar) { continue }

            $licencasTotal += $sku.prepaidUnits.enabled
            $licencasUsadas += $sku.consumedUnits

            # Traduzir nome da licenca
            $skuName = switch -Wildcard ($sku.skuPartNumber) {
                "*BUSINESS_BASIC*" { "Microsoft 365 Business Basic" }
                "*BUSINESS_STANDARD*" { "Microsoft 365 Business Standard" }
                "*BUSINESS_PREMIUM*" { "Microsoft 365 Business Premium" }
                "*O365_BUSINESS_ESSENTIALS*" { "Office 365 Business Essentials" }
                "*O365_BUSINESS_PREMIUM*" { "Office 365 Business Premium" }
                "*E1*" { "Microsoft 365 E1" }
                "*E3*" { "Microsoft 365 E3" }
                "*E5*" { "Microsoft 365 E5" }
                "*F1*" { "Microsoft 365 F1" }
                "*F3*" { "Microsoft 365 F3" }
                "*EXCHANGESTANDARD*" { "Exchange Online Plan 1" }
                "*EXCHANGEENTERPRISE*" { "Exchange Online Plan 2" }
                "*SHAREPOINTSTANDARD*" { "SharePoint Online Plan 1" }
                "*SHAREPOINTENTERPRISE*" { "SharePoint Online Plan 2" }
                "*EMS*" { "Enterprise Mobility + Security" }
                "*AAD_PREMIUM*" { "Azure AD Premium" }
                "*INTUNE*" { "Microsoft Intune" }
                "*DEFENDER*" { "Microsoft Defender" }
                "*ATP*" { "Advanced Threat Protection" }
                default { $sku.skuPartNumber }
            }

            $licencasDetalhes += [PSCustomObject]@{
                Nome       = $skuName
                Total      = $sku.prepaidUnits.enabled
                EmUso      = $sku.consumedUnits
                Disponivel = $sku.prepaidUnits.enabled - $sku.consumedUnits
            }
        }
    }
    Write-Log "Licencas: $licencasUsadas de $licencasTotal em uso" "SUCCESS"
}

# --- Reports de Uso (requer Reports.Read.All) ---
Write-Log "Coletando reports de uso..."

# OneDrive
$onedriveUsado = "N/D"
try {
    $onedriveReport = Invoke-GraphRequestBeta -Token $Token -Endpoint "/reports/getOneDriveUsageStorage(period='D30')?`$format=application/json"
    if ($onedriveReport -and $onedriveReport.value) {
        $totalBytes = ($onedriveReport.value | Measure-Object -Property storageUsedInBytes -Sum).Sum
        $onedriveUsado = "{0:N2} GB" -f ($totalBytes / 1GB)
    }
}
catch {
    Write-Log "OneDrive report nao disponivel" "WARN"
}

# SharePoint - usa Detail para pegar TODOS os sites, nao so os ativos
$sharepointUsado = "N/D"
try {
    $spReport = Invoke-GraphRequestBeta -Token $Token -Endpoint "/reports/getSharePointSiteUsageDetail(period='D180')?`$format=application/json"
    if ($spReport -and $spReport.value) {
        $totalBytes = ($spReport.value | Measure-Object -Property storageUsedInBytes -Sum).Sum
        $sharepointUsado = "{0:N2} GB" -f ($totalBytes / 1GB)
        Write-Log "SharePoint: $($spReport.value.Count) sites encontrados" "INFO"
    }
}
catch {
    Write-Log "SharePoint report nao disponivel" "WARN"
}

# Exchange
$exchangeUsado = "N/D"
try {
    $exReport = Invoke-GraphRequestBeta -Token $Token -Endpoint "/reports/getMailboxUsageStorage(period='D30')?`$format=application/json"
    if ($exReport -and $exReport.value) {
        $totalBytes = ($exReport.value | Measure-Object -Property storageUsedInBytes -Sum).Sum
        $exchangeUsado = "{0:N2} GB" -f ($totalBytes / 1GB)
    }
}
catch {
    Write-Log "Exchange report nao disponivel" "WARN"
}

# Uso por servico
Write-Log "Coletando uso por servico..."
$usoServicos = @()

try {
    $m365Report = Invoke-GraphRequestBeta -Token $Token -Endpoint "/reports/getOffice365ActiveUserCounts(period='D30')?`$format=application/json"
    if ($m365Report -and $m365Report.value) {
        $latest = $m365Report.value | Select-Object -Last 1

        $servicos = @(
            @{ Nome = "Exchange"; Ativos = $latest.exchange },
            @{ Nome = "OneDrive"; Ativos = $latest.oneDrive },
            @{ Nome = "SharePoint"; Ativos = $latest.sharePoint },
            @{ Nome = "Teams"; Ativos = $latest.teams },
            @{ Nome = "Yammer"; Ativos = $latest.yammer }
        )

        foreach ($svc in $servicos) {
            if ($null -ne $svc.Ativos) {
                $pct = if ($totalUsuarios -gt 0) { [math]::Round(($svc.Ativos / $totalUsuarios) * 100, 0) } else { 0 }
                $usoServicos += [PSCustomObject]@{
                    Servico = $svc.Nome
                    Ativos  = $svc.Ativos
                    Percent = $pct
                }
            }
        }
    }
}
catch {
    Write-Log "Reports de uso por servico nao disponiveis" "WARN"
    # Fallback com dados basicos
    $usoServicos += [PSCustomObject]@{ Servico = "Exchange"; Ativos = $totalUsuarios; Percent = 100 }
}

# --- Groups e Teams ---
Write-Log "Coletando grupos e teams..."
$totalGroups = 0
$totalTeams = 0

try {
    $groups = Invoke-GraphRequest -Token $Token -Endpoint "/groups?`$filter=groupTypes/any(c:c eq 'Unified')&`$select=id,displayName,resourceProvisioningOptions&`$top=999"
    if ($groups -and $groups.value) {
        $totalGroups = $groups.value.Count
        $totalTeams = ($groups.value | Where-Object { $_.resourceProvisioningOptions -contains "Team" }).Count
    }
    Write-Log "Grupos M365: $totalGroups, Teams: $totalTeams" "SUCCESS"
}
catch {
    Write-Log "Grupos/Teams nao disponiveis" "WARN"
}

# --- Mailboxes ---
Write-Log "Coletando mailboxes..."
$totalMailboxes = $totalUsuarios
$sharedMailboxes = 0

try {
    $mailboxReport = Invoke-GraphRequestBeta -Token $Token -Endpoint "/reports/getMailboxUsageDetail(period='D30')?`$format=application/json"
    if ($mailboxReport -and $mailboxReport.value) {
        $totalMailboxes = $mailboxReport.value.Count
    }
}
catch {
    Write-Log "Mailbox report nao disponivel" "WARN"
}

# --- Admins com MFA ---
$adminsMFA = if ($usuariosComMFA -gt 0) { "Sim" } else { "Nao" }

# --- Storage Total ---
$storageTotal = "N/D"
$storageBytes = 0

# Funcao para parsear numero no formato brasileiro ou americano
function Parse-StorageValue {
    param([string]$Value)
    if ($Value -eq "N/D" -or [string]::IsNullOrEmpty($Value)) { return 0 }
    # Remove "GB" e espacos
    $num = $Value -replace '\s*GB\s*', '' -replace '\s', ''
    # Se tem virgula como decimal (formato BR), remove pontos de milhar e troca virgula por ponto
    if ($num -match ',') {
        $num = $num -replace '\.', '' -replace ',', '.'
    }
    try { return [double]$num } catch { return 0 }
}

$storageBytes += Parse-StorageValue $onedriveUsado
$storageBytes += Parse-StorageValue $sharepointUsado
$storageBytes += Parse-StorageValue $exchangeUsado

if ($storageBytes -gt 0) {
    $storageTotal = "{0:N2} GB" -f $storageBytes
}

Write-Log "Coleta de dados finalizada" "SUCCESS"

# ============================================================================
# GERAR HTML
# ============================================================================

Write-Log "Gerando relatorio HTML..."

$templateHtml = Get-Content $TemplatePath -Raw -Encoding UTF8

# --- Carregar logo como base64 ---
$logoPath = Join-Path $ScriptPath "templates\logo.png"
$logoBase64 = ""
if (Test-Path $logoPath) {
    $logoBytes = [System.IO.File]::ReadAllBytes($logoPath)
    $logoBase64 = "data:image/png;base64," + [System.Convert]::ToBase64String($logoBytes)
    Write-Log "Logo carregado" "SUCCESS"
}
else {
    Write-Log "Logo nao encontrado: $logoPath" "WARN"
}

# Funcao helper para classes CSS
function Get-ScoreClass {
    param([int]$Score)
    if ($Score -ge 70) { return "success" }
    elseif ($Score -ge 40) { return "warning" }
    else { return "danger" }
}

function Get-BarClass {
    param([int]$Score)
    if ($Score -ge 70) { return "" }
    elseif ($Score -ge 40) { return "medium" }
    else { return "low" }
}

# Substituicoes
$html = $templateHtml

# Logo
$html = $html -replace "{{LOGO_BASE64}}", $logoBase64

# Basicos
$html = $html -replace "{{CLIENTE_NOME}}", $ClienteNome
$html = $html -replace "{{PERIODO}}", $periodo
$html = $html -replace "{{SECURE_SCORE}}", $secureScore
$html = $html -replace "{{SECURE_SCORE_CLASS}}", (Get-ScoreClass $secureScore)
$html = $html -replace "{{SECURE_SCORE_BAR_CLASS}}", (Get-BarClass $secureScore)
$html = $html -replace "{{SECURE_SCORE_ATUAL}}", $secureScoreAtual
$html = $html -replace "{{SECURE_SCORE_MAX}}", $secureScoreMax
$html = $html -replace "{{SECURE_SCORE_VARIACAO}}", $secureScoreVariacao

$variacaoClass = if ($secureScoreVariacao -like "+*") { "good" } elseif ($secureScoreVariacao -like "-*") { "bad" } else { "" }
$html = $html -replace "{{SECURE_SCORE_VARIACAO_CLASS}}", $variacaoClass

$html = $html -replace "{{TOTAL_USUARIOS}}", $totalUsuarios
$html = $html -replace "{{MFA_PERCENT}}", $mfaPercent
$html = $html -replace "{{MFA_CLASS}}", (Get-ScoreClass $mfaPercent)
$html = $html -replace "{{USUARIOS_MFA}}", $usuariosComMFA

# MFA Value Class
$mfaValueClass = if ($mfaPercent -ge 80) { "good" } elseif ($mfaPercent -ge 50) { "warning" } else { "bad" }
$html = $html -replace "{{MFA_VALUE_CLASS}}", $mfaValueClass

# Admins MFA
$html = $html -replace "{{ADMINS_MFA}}", $adminsMFA

# Executive Summary
$execSecurityClass = if ($secureScore -ge 50) { "green" } elseif ($secureScore -ge 30) { "yellow" } else { "yellow" }
$execSecurityIcon = if ($secureScore -ge 50) { [char]0x2713 } else { "!" }
$execSecurityText = if ($secureScore -ge 50) { "Seguranca adequada" } else { "Seguranca requer atencao" }
$html = $html -replace [regex]::Escape("{{EXEC_SECURITY_CLASS}}"), $execSecurityClass
$html = $html -replace [regex]::Escape("{{EXEC_SECURITY_ICON}}"), $execSecurityIcon
$html = $html -replace [regex]::Escape("{{EXEC_SECURITY_TEXT}}"), $execSecurityText

# Security Analysis (sempre verde no Basic - nao tem ATP)
$html = $html -replace [regex]::Escape("{{THREAT_ICON}}"), '<span style="color:#00FF00">0</span>'
$html = $html -replace [regex]::Escape("{{MALWARE_ICON}}"), '<span style="color:#00FF00">0</span>'
$html = $html -replace [regex]::Escape("{{PHISHING_ICON}}"), '<span style="color:#00FF00">0</span>'
$html = $html -replace [regex]::Escape("{{SECURITY_STATUS}}"), "Nenhum incidente detectado"

# Colaboracao
$html = $html -replace "{{TOTAL_TEAMS}}", $totalTeams
$html = $html -replace "{{TOTAL_GROUPS}}", $totalGroups
$html = $html -replace "{{TOTAL_MAILBOXES}}", $totalMailboxes
$html = $html -replace "{{SHARED_MAILBOXES}}", $sharedMailboxes

$html = $html -replace "{{TOTAL_LOGINS}}", $(if ($signInSuccessEstimado -eq "N/D") { "N/D" } else { $signInSuccessEstimado + $signInFailedEstimado })
$html = $html -replace "{{LOGINS_SUCESSO}}", $signInSuccessEstimado
$html = $html -replace "{{LOGINS_FALHOS}}", $signInFailedEstimado

# Classes para logins - tratar "N/D" como info (azul)
if ($signInFailedEstimado -eq "N/D") {
    $loginsClass = "info"
    $loginsValueClass = ""
}
elseif ($signInFailedEstimado -gt 100) {
    $loginsClass = "danger"
    $loginsValueClass = "bad"
}
elseif ($signInFailedEstimado -gt 50) {
    $loginsClass = "warning"
    $loginsValueClass = "warning"
}
else {
    $loginsClass = "success"
    $loginsValueClass = ""
}
$html = $html -replace "{{LOGINS_FALHOS_CLASS}}", $loginsClass
$html = $html -replace "{{LOGINS_FALHOS_VALUE_CLASS}}", $loginsValueClass

# Usuarios sem MFA
$usuariosSemMFASection = ""
if ($Config.opcoes.incluir_usuarios_sem_mfa -and $usuariosSemMFA.Count -gt 0) {
    $usuariosSemMFASection = @"
<div class="alert-box">
    <h4>Usuarios sem MFA ($($usuariosSemMFA.Count))</h4>
    <ul>
"@
    foreach ($u in $usuariosSemMFA) {
        $usuariosSemMFASection += "<li>$u</li>`n"
    }
    $usuariosSemMFASection += "</ul></div>"
}
$html = $html -replace "{{USUARIOS_SEM_MFA_SECTION}}", $usuariosSemMFASection

# Licencas
$html = $html -replace "{{LICENCAS_TOTAL}}", $licencasTotal
$html = $html -replace "{{LICENCAS_USADAS}}", $licencasUsadas
$html = $html -replace "{{LICENCAS_DISPONIVEIS}}", ($licencasTotal - $licencasUsadas)

# Taxa de utilizacao
$licencasUtilizacao = if ($licencasTotal -gt 0) { [math]::Round(($licencasUsadas / $licencasTotal) * 100, 0) } else { 0 }
$html = $html -replace "{{LICENCAS_UTILIZACAO}}", $licencasUtilizacao

$licencasDetalhesHtml = ""
if ($licencasDetalhes.Count -gt 0) {
    $licencasDetalhesHtml = "<table style='margin-top: 15px;'><thead><tr><th>Plano</th><th>Total</th><th>Em Uso</th><th>Disponivel</th></tr></thead><tbody>"
    foreach ($lic in $licencasDetalhes) {
        $licencasDetalhesHtml += "<tr><td>$($lic.Nome)</td><td>$($lic.Total)</td><td>$($lic.EmUso)</td><td>$($lic.Disponivel)</td></tr>"
    }
    $licencasDetalhesHtml += "</tbody></table>"
}
$html = $html -replace "{{LICENCAS_DETALHES}}", $licencasDetalhesHtml

# Storage
$html = $html -replace "{{ONEDRIVE_USADO}}", $onedriveUsado
$html = $html -replace "{{SHAREPOINT_USADO}}", $sharepointUsado
$html = $html -replace "{{EXCHANGE_USADO}}", $exchangeUsado
$html = $html -replace "{{STORAGE_TOTAL}}", $storageTotal

# Tabela de uso de servicos
$tabelaUsoServicos = ""
foreach ($svc in $usoServicos) {
    $tabelaUsoServicos += "<tr><td>$($svc.Servico)</td><td>$($svc.Ativos)</td><td>$($svc.Percent)%</td></tr>`n"
}
if ($tabelaUsoServicos -eq "") {
    $tabelaUsoServicos = "<tr><td colspan='3' style='text-align:center; color:#A0A0A0;'>Dados de uso nao disponiveis</td></tr>"
}
$html = $html -replace "{{TABELA_USO_SERVICOS}}", $tabelaUsoServicos

# Recomendacoes
$recomendacoes = @()
if ($mfaPercent -lt 100) { $recomendacoes += "Habilitar MFA para todos os usuarios (atualmente $mfaPercent%)" }
if ($secureScore -lt 50) { $recomendacoes += "Melhorar Secure Score - revisar recomendacoes no portal Microsoft 365" }
if ($signInFailedEstimado -ne "N/D" -and $signInFailedEstimado -gt 50) { $recomendacoes += "Investigar alto numero de tentativas de login falhas" }
if ($signInFailedEstimado -eq "N/D") { $recomendacoes += "Considerar Azure AD Premium P1 para monitoramento detalhado de logins" }

$recomendacoesSection = ""
if ($Config.opcoes.incluir_recomendacoes -and $recomendacoes.Count -gt 0) {
    $recomendacoesSection = @"
<div class="section">
    <h2 class="section-title">Recomendacoes de Seguranca</h2>
    <div class="recommendations-box">
        <h4>Acoes Sugeridas</h4>
        <ul>
"@
    foreach ($rec in $recomendacoes) {
        $recomendacoesSection += "<li>$rec</li>`n"
    }
    $recomendacoesSection += "</ul></div></div>"
}
$html = $html -replace "{{RECOMENDACOES_SECTION}}", $recomendacoesSection

# Resumo
$resumo = "Ambiente Microsoft 365 monitorado. "
if ($secureScore -ge 70) { $resumo += "Postura de seguranca adequada. " }
elseif ($secureScore -ge 40) { $resumo += "Postura de seguranca pode ser melhorada. " }
else { $resumo += "Atencao: postura de seguranca requer melhorias urgentes. " }

if ($mfaPercent -eq 100) { $resumo += "Todos os usuarios com MFA habilitado. " }
elseif ($mfaPercent -ge 80) { $resumo += "Maioria dos usuarios protegidos com MFA. " }

$resumo += "Nenhum incidente critico detectado no periodo."
$html = $html -replace "{{RESUMO_PERIODO}}", $resumo

# MSP info
$html = $html -replace "{{MSP_NOME}}", $Config.msp.nome
$html = $html -replace "{{MSP_CONTATO}}", $Config.msp.contato
$html = $html -replace "{{DATA_GERACAO}}", (Get-Date -Format "dd/MM/yyyy HH:mm")

# ============================================================================
# SALVAR HTML E GERAR PDF
# ============================================================================

$timestamp = Get-Date -Format "yyyyMM"
$nomeArquivo = "$($ClienteNome -replace ' ', '_')_M365_Report_$timestamp"
$htmlFile = Join-Path $OutputPath "$nomeArquivo.html"
$pdfFile = Join-Path $OutputPath "$nomeArquivo.pdf"

# Salvar HTML
$html | Out-File -FilePath $htmlFile -Encoding UTF8
Write-Log "HTML salvo: $htmlFile" "SUCCESS"

if ($Preview) {
    Write-Log "Abrindo preview no navegador..."
    Start-Process $htmlFile
    Write-Log "Preview aberto. PDF nao gerado (modo preview)." "SUCCESS"
    exit 0
}

# Gerar PDF usando Edge
Write-Log "Gerando PDF com Microsoft Edge..."

$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (!(Test-Path $edgePath)) {
    $edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
}

if (Test-Path $edgePath) {
    $htmlFullPath = (Resolve-Path $htmlFile).Path
    $pdfFullPath = Join-Path (Resolve-Path $OutputPath).Path "$nomeArquivo.pdf"

    # Edge headless print to PDF
    $arguments = "--headless --disable-gpu --print-to-pdf=`"$pdfFullPath`" --no-pdf-header-footer `"file:///$($htmlFullPath -replace '\\', '/')`""

    Start-Process -FilePath $edgePath -ArgumentList $arguments -Wait -NoNewWindow

    Start-Sleep -Seconds 2

    if (Test-Path $pdfFullPath) {
        Write-Log "PDF gerado: $pdfFullPath" "SUCCESS"
        $pdfFile = $pdfFullPath
    }
    else {
        Write-Log "Falha ao gerar PDF. Verifique o Edge." "ERROR"
        exit 1
    }
}
else {
    Write-Log "Microsoft Edge nao encontrado. Instale o Edge ou use -Preview" "ERROR"
    exit 1
}

# ============================================================================
# ENVIAR EMAIL (OPCIONAL)
# ============================================================================

if ($SendEmail) {
    Write-Log "Preparando envio de email..."

    # Ler PDF como base64
    $pdfBytes = [System.IO.File]::ReadAllBytes($pdfFile)
    $pdfBase64 = [System.Convert]::ToBase64String($pdfBytes)

    # Montar corpo do email
    $assunto = $Config.email.assunto_template -replace "{{PERIODO}}", $periodo

    $emailBody = @"
<html>
<body style="font-family: Segoe UI, sans-serif; color: #333;">
<p>Prezado(a),</p>

<p>Segue em anexo o <strong>Relatorio Mensal de Saude do Microsoft 365</strong> referente a <strong>$periodo</strong>.</p>

<p><strong>Resumo:</strong></p>
<ul>
    <li>Secure Score: $secureScore%</li>
    <li>Usuarios Ativos: $totalUsuarios</li>
    <li>MFA Habilitado: $mfaPercent%</li>
</ul>

<p>Em caso de duvidas, entre em contato conosco.</p>

<p>Atenciosamente,<br>
<strong>$($Config.msp.nome)</strong><br>
$($Config.msp.contato)</p>
</body>
</html>
"@

    # Criar mensagem via Graph
    $message = @{
        message = @{
            subject = $assunto
            body    = @{
                contentType = "HTML"
                content     = $emailBody
            }
            toRecipients = @(
                $Config.email.destinatarios | ForEach-Object {
                    @{ emailAddress = @{ address = $_ } }
                }
            )
            attachments = @(
                @{
                    "@odata.type"  = "#microsoft.graph.fileAttachment"
                    name           = "$nomeArquivo.pdf"
                    contentType    = "application/pdf"
                    contentBytes   = $pdfBase64
                }
            )
        }
        saveToSentItems = $true
    }

    # Adicionar CC se houver
    if ($Config.email.cc -and $Config.email.cc.Count -gt 0) {
        $message.message.ccRecipients = @(
            $Config.email.cc | ForEach-Object {
                @{ emailAddress = @{ address = $_ } }
            }
        )
    }

    # Enviar via Graph (shared mailbox)
    $sendEndpoint = "/users/$($Config.msp.remetente)/sendMail"

    try {
        $headers = @{
            "Authorization" = "Bearer $Token"
            "Content-Type"  = "application/json"
        }

        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0$sendEndpoint" -Method Post -Headers $headers -Body ($message | ConvertTo-Json -Depth 10)

        Write-Log "Email enviado com sucesso para: $($Config.email.destinatarios -join ', ')" "SUCCESS"
    }
    catch {
        Write-Log "Falha ao enviar email: $_" "ERROR"
        Write-Log "O PDF foi gerado em: $pdfFile" "WARN"
    }
}

# ============================================================================
# FINALIZAR
# ============================================================================

Write-Log "========================================"
Write-Log "RELATORIO CONCLUIDO"
Write-Log "Cliente: $ClienteNome"
Write-Log "Periodo: $periodo"
Write-Log "PDF: $pdfFile"
Write-Log "Log: $LogFile"
Write-Log "========================================"

# Retornar caminho do PDF
return $pdfFile
