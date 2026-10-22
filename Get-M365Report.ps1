<#
.SYNOPSIS
    Collects Microsoft 365 data through the Microsoft Graph REST API and builds a
    monthly health/security report (HTML + PDF).

.DESCRIPTION
    Monthly Microsoft 365 health report for MSPs. Uses raw Graph REST calls with an
    app registration (client credentials). No PowerShell modules are required.

    Collected (each source is optional; a missing permission/license renders
    "Not collected" instead of a made-up number):
      - Secure Score (current and ~30 days ago)
      - Enabled member users, MFA registration (per user and per administrator)
      - Sign-in logs, last 30 days (requires Entra ID P1/P2)
      - Security alerts, last 30 days (alerts_v2)
      - Service health (current status + incidents/advisories, last 30 days)
      - Licenses, Microsoft 365 groups/Teams
      - Usage reports: storage (OneDrive/SharePoint/Exchange), mailboxes,
        active users per service

    Required Graph permissions (Application, admin consent) on the CLIENT app:
      User.Read.All                      users
      Directory.Read.All                 admin roles, groups, domains, licenses
      AuditLog.Read.All                  sign-in logs, MFA registration report
      Reports.Read.All                   usage reports
      SecurityEvents.Read.All            Secure Score
      SecurityAlert.Read.All             security alerts
      ServiceHealth.Read.All             service health
      UserAuthenticationMethod.Read.All  MFA fallback (only if the registration report fails)

    For -SendEmail, the sender app in the MSP tenant (msp.sender_app) needs
    Mail.Send (Application). Without a sender app, the client app needs Mail.Send
    and the sender mailbox must belong to the client tenant.
    Any permission missing renders that section as "Not collected".

.PARAMETER ConfigFile
    Path to the client JSON configuration file.

.PARAMETER OutputPath
    Folder where the HTML/PDF are saved. Default: ./output

.PARAMETER Preview
    Opens the HTML in the browser instead of generating a PDF.

.PARAMETER SendEmail
    Sends the PDF by email after it is generated.

.EXAMPLE
    .\Get-M365Report.ps1 -ConfigFile ".\config\contoso.json" -Preview

.EXAMPLE
    .\Get-M365Report.ps1 -ConfigFile ".\config\contoso.json" -SendEmail
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
# SETUP
# ============================================================================

$ErrorActionPreference = "Stop"
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$TemplatePath = Join-Path (Join-Path $ScriptPath "templates") "report.html"
$LangPath = Join-Path $ScriptPath "lang"
$LogPath = Join-Path $ScriptPath "logs"

if (!(Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
if (!(Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }

$LogFile = Join-Path $LogPath "report_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $logMessage

    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARN" { Write-Host $logMessage -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage }
    }
}

# Reads the first non-empty value among several dotted config paths.
# English keys are listed first; legacy Portuguese keys remain supported.
function Get-ConfigValue {
    param($Object, [string[]]$Path, $Default = $null)
    foreach ($p in $Path) {
        $cur = $Object
        foreach ($seg in $p.Split('.')) {
            if ($null -eq $cur) { break }
            $cur = $cur.$seg
        }
        if ($null -ne $cur -and -not ($cur -is [string] -and $cur.Trim() -eq '')) { return $cur }
    }
    return $Default
}

# ============================================================================
# LOAD CONFIGURATION
# ============================================================================

Write-Log "Starting Microsoft 365 data collection..."
Write-Log "Config file: $ConfigFile"

if (!(Test-Path $ConfigFile)) {
    Write-Log "Configuration file not found: $ConfigFile" "ERROR"
    exit 1
}

$ConfigFile = (Resolve-Path $ConfigFile).Path
$ConfigDir = Split-Path -Parent $ConfigFile
$Config = Get-Content $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json

$ClientName = [string](Get-ConfigValue $Config @('client.name', 'cliente.nome') '')
$TenantId = Get-ConfigValue $Config @('client.tenant_id', 'cliente.tenant_id')
$ClientId = Get-ConfigValue $Config @('client.client_id', 'cliente.client_id')
$ClientSecret = Get-ConfigValue $Config @('client.client_secret', 'cliente.client_secret')
$ClientCertThumbprint = Get-ConfigValue $Config @('client.certificate_thumbprint', 'cliente.certificate_thumbprint')

$MspName = [string](Get-ConfigValue $Config @('msp.name', 'msp.nome') '')
$MspContact = [string](Get-ConfigValue $Config @('msp.contact', 'msp.contato') '')
$MspSender = [string](Get-ConfigValue $Config @('msp.sender', 'msp.remetente') '')
$MspTagline = [string](Get-ConfigValue $Config @('msp.tagline', 'msp.slogan') '')
$MspLogo = Get-ConfigValue $Config @('msp.logo')

$IncludeUsersWithoutMfa = [bool](Get-ConfigValue $Config @('options.include_users_without_mfa', 'opcoes.incluir_usuarios_sem_mfa') $true)
$IncludeRecommendations = [bool](Get-ConfigValue $Config @('options.include_recommendations', 'opcoes.incluir_recomendacoes') $true)
$Language = [string](Get-ConfigValue $Config @('options.language', 'opcoes.idioma') 'en')
$MaxSignInPages = [int](Get-ConfigValue $Config @('options.max_signin_pages') 50)

if (-not $TenantId -or -not $ClientId -or (-not $ClientSecret -and -not $ClientCertThumbprint)) {
    Write-Log "Config must define client tenant_id, client_id and either client_secret or certificate_thumbprint." "ERROR"
    exit 1
}

Write-Log "Client: $ClientName"
Write-Log "Tenant: $TenantId"

# ============================================================================
# STRINGS (English built in; other languages loaded from lang\<code>.json)
# ============================================================================

$Strings = @{
    # Template labels
    TITLE              = "Microsoft 365 Report"
    HEADER_TITLE       = "Health Report"
    EXEC_SUMMARY       = "Executive Summary"
    SERVICE_HEALTH     = "Microsoft 365 service health"
    INCIDENTS_30D      = "Microsoft incidents (30 days)"
    HEALTH_SOURCE      = "Source: Microsoft Graph service health API"
    USERS              = "Users"
    MFA_REGISTERED     = "MFA Registered"
    FAILED_SIGNINS     = "Failed sign-ins (30d)"
    SECURITY_TITLE     = "Security Alerts"
    SECURITY_DESC      = "Alerts raised by Microsoft security products in the last 30 days (Microsoft Graph security alerts API)."
    ALERTS_TOTAL       = "Total alerts"
    ALERTS_HIGH        = "High severity"
    ALERTS_ACTIVE      = "Unresolved"
    SECURITY_STATUS    = "Overall security status"
    SECURE_SCORE_DESC  = "Microsoft's official metric for the security posture of the tenant."
    CURRENT_SCORE      = "Current score"
    MAX_SCORE          = "Maximum possible score"
    SCORE_CHANGE       = "Change vs. ~30 days ago"
    AUTH_TITLE         = "Authentication & Identity"
    SIGNINS_PERIOD     = "Sign-ins (last 30 days)"
    TOTAL_SIGNINS      = "Total sign-ins"
    SUCCESSFUL         = "Successful"
    FAILED_INTERRUPTED = "Failed or interrupted"
    ACCOUNT_PROTECTION = "Account Protection"
    USERS_WITH_MFA     = "Users with MFA registered"
    MFA_COVERAGE       = "MFA coverage"
    ADMINS_WITH_MFA    = "Admins with MFA registered"
    COLLAB_TITLE       = "Collaboration & Productivity"
    TOTAL_TEAMS        = "Teams"
    M365_GROUPS        = "Microsoft 365 groups"
    MAIL               = "Email"
    MAILBOXES          = "Mailboxes"
    SHARED_MAILBOXES   = "Shared mailboxes"
    LICENSING          = "Licensing"
    LIC_PURCHASED      = "Licenses purchased"
    LIC_ASSIGNED       = "Licenses assigned"
    LIC_AVAILABLE      = "Licenses available"
    LIC_UTILIZATION    = "Utilization"
    STORAGE_TITLE      = "Cloud Storage"
    STORAGE_DESC       = "Most recent daily snapshot from Microsoft 365 usage reports."
    TOTAL_STORED       = "Total stored"
    ADOPTION_TITLE     = "Service Adoption"
    ADOPTION_DESC      = "Active users per Microsoft 365 service in the last 30 days."
    SERVICE            = "Service"
    ACTIVE_USERS       = "Active users"
    ADOPTION           = "Adoption"
    TECH_SUMMARY       = "Technical Summary"
    GENERATED_AT       = "Report generated automatically on"
    DATA_SOURCE        = "Data collected via Microsoft Graph API"

    # Script-generated text
    NotCollected       = "Not collected"
    Points             = "{0} points"
    OfFormat           = "{0} of {1}"
    Partial            = "{0} (partial)"
    ExecSourcesAll     = "All {0} data sources collected"
    ExecSourcesPartial = "{0} of {1} data sources collected"
    ExecUsers          = "{0} active user accounts"
    ExecUsersUnknown   = "Active user count not collected"
    SecurityAdequate   = "Security posture adequate"
    SecurityAttention  = "Security posture needs attention"
    SecurityUnknown    = "Secure Score not collected"
    HealthAllOk        = "All {0} monitored services operational at report time"
    HealthIssues       = "{0} of {1} services reporting issues at report time"
    HealthNotCollected = "Service health not collected (requires ServiceHealth.Read.All)"
    HealthIncidentsFmt = "{0} incident(s), {1} advisory(ies)"
    AlertsNone         = "No security alerts in the last 30 days"
    AlertsReview       = "{0} unresolved alert(s) need review"
    AlertsAllResolved  = "{0} alert(s), all resolved"
    AlertsNotCollected = "Security alerts not collected (requires SecurityAlert.Read.All)"
    SigninsNotCollected = "Sign-in logs not available (requires Entra ID P1/P2 and AuditLog.Read.All)."
    SigninsTruncated   = "Partial count: only the most recent {0} sign-in events were read."
    SigninsNote        = "Interrupted sign-ins (e.g. MFA prompts) are counted as failed."
    UsersWithoutMfa    = "Users without MFA registered ({0})"
    AdminsWithoutMfa   = "Administrators without MFA registered ({0})"
    NoAdminsFound      = "No admin role members found"
    LicPlan            = "Plan"
    LicTotal           = "Total"
    LicAssigned        = "Assigned"
    LicAvailable       = "Available"
    UsageNotAvailable  = "Usage data not available (requires Reports.Read.All)"
    RecTitle           = "Security Recommendations"
    RecSubtitle        = "Suggested actions"
    RecMfa             = "Require MFA registration for all users (currently {0}%)."
    RecAdminMfa        = "Register MFA for every administrator account ({0} without MFA)."
    RecSecureScore     = "Improve Secure Score - review the improvement actions in the Microsoft Defender portal."
    RecFailedSignins   = "Investigate the high number of failed sign-ins ({0} in the last 30 days)."
    RecSigninsP1       = "Sign-in logs could not be read. They require an Entra ID P1/P2 license and the AuditLog.Read.All permission."
    RecAlerts          = "Review the {0} unresolved security alert(s) in the Microsoft Defender portal."
    SumScoreGood       = "Secure Score is {0}%, an adequate security posture."
    SumScoreMid        = "Secure Score is {0}%; the security posture can be improved."
    SumScoreLow        = "Secure Score is {0}%; the security posture needs urgent improvement."
    SumScoreNA         = "Secure Score could not be collected."
    SumMfa             = "{0}% of active users have MFA registered."
    SumMfaNA           = "MFA registration status could not be collected."
    SumAdminsOk        = "All {0} administrator account(s) have MFA registered."
    SumAdminsBad       = "{0} of {1} administrator account(s) have no MFA registered."
    SumAlertsNone      = "No security alerts were raised in the last 30 days."
    SumAlerts          = "{0} security alert(s) were raised in the last 30 days ({1} high severity, {2} unresolved)."
    SumAlertsNA        = "Security alerts were not collected for this report."
    SumIncidents       = "Microsoft reported {0} service incident(s) in the last 30 days."
    SumSignins         = "{0} failed or interrupted sign-ins were recorded in the last 30 days."
    EmailSubject       = "Microsoft 365 Monthly Report - {{PERIOD}}"
    EmailGreeting      = "Hello,"
    EmailIntro         = "Please find attached the <strong>Microsoft 365 Monthly Health Report</strong> for <strong>{0}</strong>."
    EmailSummary       = "Summary:"
    EmailUsers         = "Active users"
    EmailMfa           = "MFA registered"
    EmailQuestions     = "If you have any questions, please contact us."
    EmailRegards       = "Best regards,"
    DateFormat         = "yyyy-MM-dd HH:mm"
}

$CultureName = "en-US"
if ($Language -and $Language -notlike "en*") {
    $langFile = Join-Path $LangPath "$Language.json"
    if (Test-Path $langFile) {
        $langData = Get-Content $langFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($prop in $langData.PSObject.Properties) { $Strings[$prop.Name] = [string]$prop.Value }
        $CultureName = $Language
        Write-Log "Language: $Language"
    }
    else {
        Write-Log "Language file not found ($langFile). Falling back to English." "WARN"
        $Language = "en"
    }
}
else {
    $Language = "en"
}
try { $Culture = [System.Globalization.CultureInfo]::GetCultureInfo($CultureName) }
catch { $Culture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US") }

function T {
    param([string]$Key, [object[]]$FormatArgs)
    $s = $Strings[$Key]
    if ($null -eq $s) { $s = $Key }
    if ($PSBoundParameters.ContainsKey('FormatArgs')) { return [string]::Format($Culture, $s, $FormatArgs) }
    return $s
}

function Format-Num {
    param($Value)
    return [string]::Format($Culture, "{0:N0}", [double]$Value)
}

function Format-Bytes {
    param([double]$Bytes)
    return [string]::Format($Culture, "{0:N2} GB", ($Bytes / 1GB))
}

function Enc {
    param($Value)
    if ($null -eq $Value) { return "" }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

# ============================================================================
# GRAPH AUTHENTICATION
# ============================================================================

function ConvertTo-Base64Url {
    param([byte[]]$Bytes)
    return [System.Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Get-CertificateFromStore {
    param([string]$Thumbprint)
    $thumb = ($Thumbprint -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
    foreach ($location in @("CurrentUser", "LocalMachine")) {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", $location)
        try {
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
            $found = $store.Certificates.Find([System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $thumb, $false)
            foreach ($cert in $found) { if ($cert.HasPrivateKey) { return $cert } }
        }
        catch { }
        finally { $store.Close() }
    }
    throw "Certificate with thumbprint $thumb and a private key was not found in CurrentUser\My or LocalMachine\My."
}

function Get-GraphToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$CertificateThumbprint
    )

    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        grant_type = "client_credentials"
        client_id  = $ClientId
        scope      = "https://graph.microsoft.com/.default"
    }

    if ($CertificateThumbprint) {
        # Certificate auth: signed JWT client assertion (RS256)
        $cert = Get-CertificateFromStore -Thumbprint $CertificateThumbprint
        $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $header = [ordered]@{ alg = "RS256"; typ = "JWT"; x5t = (ConvertTo-Base64Url $cert.GetCertHash()) }
        $payload = [ordered]@{
            aud = $tokenUrl
            iss = $ClientId
            sub = $ClientId
            jti = [guid]::NewGuid().ToString()
            nbf = $now - 60
            exp = $now + 600
        }
        $enc = [System.Text.Encoding]::UTF8
        $unsigned = (ConvertTo-Base64Url $enc.GetBytes(($header | ConvertTo-Json -Compress))) + "." +
                    (ConvertTo-Base64Url $enc.GetBytes(($payload | ConvertTo-Json -Compress)))
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        $sig = $rsa.SignData($enc.GetBytes($unsigned), [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $body.client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        $body.client_assertion = $unsigned + "." + (ConvertTo-Base64Url $sig)
    }
    else {
        $body.client_secret = $ClientSecret
    }

    try {
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        Write-Log "Failed to obtain token for tenant ${TenantId}: $_" "ERROR"
        throw
    }
}

# ============================================================================
# GRAPH REQUESTS
# ============================================================================

$script:LastGraphStatus = $null

function Invoke-GraphRequest {
    param(
        [string]$Token,
        [string]$Endpoint,
        [string]$Method = "GET",
        $Body = $null,
        [switch]$Beta,
        [switch]$Quiet
    )

    if ($Endpoint -match '^https://') { $uri = $Endpoint }
    elseif ($Beta) { $uri = "https://graph.microsoft.com/beta$Endpoint" }
    else { $uri = "https://graph.microsoft.com/v1.0$Endpoint" }

    $headers = @{ "Authorization" = "Bearer $Token" }
    $script:LastGraphStatus = $null

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            if ($null -ne $Body) {
                $json = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
                $response = Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers -Body $bytes -ContentType "application/json; charset=utf-8"
            }
            else {
                $response = Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers
            }
            $script:LastGraphStatus = 200
            return $response
        }
        catch {
            $status = $null
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }
            $script:LastGraphStatus = $status
            if ($status -in @(429, 503, 504) -and $attempt -lt 3) {
                Start-Sleep -Seconds (5 * $attempt)
                continue
            }
            $detail = if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
            if (-not $Quiet) { Write-Log "Graph request failed [$status] ($Endpoint): $detail" "WARN" }
            return $null
        }
    }
    return $null
}

# Follows @odata.nextLink. Returns @{ Items; Truncated } or $null on failure.
function Get-GraphCollection {
    param(
        [string]$Token,
        [string]$Endpoint,
        [switch]$Beta,
        [int]$MaxPages = 0,
        [switch]$Quiet
    )

    $items = New-Object System.Collections.Generic.List[object]
    $next = $Endpoint
    $pages = 0
    $truncated = $false

    while ($next) {
        $resp = Invoke-GraphRequest -Token $Token -Endpoint $next -Beta:$Beta -Quiet:$Quiet
        if ($null -eq $resp) {
            if ($pages -eq 0) { return $null }
            Write-Log "Paging stopped early for $Endpoint after $pages page(s)" "WARN"
            $truncated = $true
            break
        }
        foreach ($i in @($resp.value)) { if ($null -ne $i) { $items.Add($i) } }
        $pages++
        $next = $resp.'@odata.nextLink'
        if ($next -and $MaxPages -gt 0 -and $pages -ge $MaxPages) {
            $truncated = $true
            break
        }
    }

    return [PSCustomObject]@{ Items = $items.ToArray(); Truncated = $truncated }
}

# Usage reports. v1.0 returns CSV (via a redirect to a pre-authenticated URL);
# the JSON format ($format=application/json) only exists on /beta, so beta is
# used only as a fallback when the v1.0 CSV cannot be read.
function Get-UsageReport {
    param([string]$Token, [string]$Function)

    $resp = Invoke-GraphRequest -Token $Token -Endpoint "/reports/$Function" -Quiet
    if ($resp -is [string] -or $resp -is [byte[]]) {
        try {
            $text = if ($resp -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($resp) } else { $resp }
            $text = $text.TrimStart([char]0xFEFF)
            if ($text.StartsWith([string]([char]0xEF) + [char]0xBB + [char]0xBF)) { $text = $text.Substring(3) }
            if ($text.Contains('Report Refresh Date')) {
                return ,@($text | ConvertFrom-Csv)
            }
        }
        catch { }
    }

    $status = $script:LastGraphStatus
    if ($status -eq 403 -or $status -eq 401) {
        Write-Log "Usage report $Function denied [$status] - requires Reports.Read.All" "WARN"
        return $null
    }

    Write-Log "Usage report $Function not readable on v1.0 - trying beta JSON" "WARN"
    $sep = if ($Function.Contains('?')) { '&' } else { '?' }
    $beta = Get-GraphCollection -Token $Token -Endpoint "/reports/$Function$sep`$format=application/json" -Beta
    if ($null -eq $beta) { return $null }
    return ,$beta.Items
}

function Get-Field {
    param($Row, [string[]]$Names)
    foreach ($n in $Names) {
        $p = $Row.PSObject.Properties[$n]
        if ($p -and $null -ne $p.Value -and [string]$p.Value -ne '') { return $p.Value }
    }
    return $null
}

function ConvertTo-Double {
    param($Value)
    if ($null -eq $Value) { return $null }
    $d = 0.0
    if ([double]::TryParse([string]$Value, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    return $null
}

function Test-Truthy {
    param($Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return $Value }
    return ([string]$Value).Trim() -eq 'True'
}

# Latest-day storage value from a per-day storage report (rows are one per day).
function Get-LatestStorageBytes {
    param($Rows)
    if ($null -eq $Rows -or @($Rows).Count -eq 0) { return $null }
    $dated = @($Rows | ForEach-Object {
            [PSCustomObject]@{
                Date  = [string](Get-Field $_ @('Report Date', 'reportDate'))
                Type  = [string](Get-Field $_ @('Site Type', 'siteType'))
                Bytes = ConvertTo-Double (Get-Field $_ @('Storage Used (Byte)', 'storageUsedInBytes'))
            }
        } | Where-Object { $null -ne $_.Bytes -and $_.Date })
    if ($dated.Count -eq 0) { return $null }
    $latestDate = ($dated | Sort-Object Date | Select-Object -Last 1).Date
    $latest = @($dated | Where-Object { $_.Date -eq $latestDate })
    $all = @($latest | Where-Object { $_.Type -eq 'All' })
    if ($all.Count -gt 0) { return [double]$all[0].Bytes }
    return [double](($latest | Measure-Object -Property Bytes -Sum).Sum)
}

# ============================================================================
# DATA COLLECTION
# ============================================================================

$Sources = [ordered]@{}
function Set-Source { param([string]$Name, [bool]$Ok) $Sources[$Name] = $Ok }

Write-Log "Authenticating to Microsoft Graph (client tenant)..."
$Token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $ClientCertThumbprint
Write-Log "Token acquired" "SUCCESS"

$now = Get-Date
$previousMonth = $now.AddMonths(-1)
$period = $Culture.TextInfo.ToTitleCase($previousMonth.ToString("MMMM yyyy", $Culture))
$since30 = $now.ToUniversalTime().AddDays(-30)
$since30Iso = $since30.ToString("yyyy-MM-ddTHH:mm:ssZ")

Write-Log "Report period: $period"

# --- Secure Score -----------------------------------------------------------
Write-Log "Collecting Secure Score..."
$secureScore = $null
$secureScoreCurrent = $null
$secureScoreMax = $null
$secureScoreChange = $null

$ssData = Invoke-GraphRequest -Token $Token -Endpoint "/security/secureScores?`$top=40"
if ($ssData -and @($ssData.value).Count -gt 0) {
    $scores = @($ssData.value | Sort-Object { [datetime]$_.createdDateTime } -Descending)
    $latest = $scores[0]
    if ($latest.maxScore -gt 0) {
        $secureScoreCurrent = [math]::Round([double]$latest.currentScore, 1)
        $secureScoreMax = [math]::Round([double]$latest.maxScore, 1)
        $secureScore = [int][math]::Round(([double]$latest.currentScore / [double]$latest.maxScore) * 100, 0)

        $refDate = ([datetime]$latest.createdDateTime).AddDays(-28)
        $older = @($scores | Where-Object { [datetime]$_.createdDateTime -le $refDate -and $_.maxScore -gt 0 }) | Select-Object -First 1
        if ($older) {
            $olderPct = [int][math]::Round(([double]$older.currentScore / [double]$older.maxScore) * 100, 0)
            $secureScoreChange = $secureScore - $olderPct
        }
        Write-Log "Secure Score: $secureScore% ($secureScoreCurrent/$secureScoreMax)" "SUCCESS"
    }
}
Set-Source "SecureScore" ($null -ne $secureScore)
if ($null -eq $secureScore) { Write-Log "Secure Score not available" "WARN" }

# --- Users ------------------------------------------------------------------
Write-Log "Collecting users..."
$users = @()
$totalUsers = $null
$usersResult = Get-GraphCollection -Token $Token -Endpoint "/users?`$filter=accountEnabled eq true&`$select=id,displayName,userPrincipalName,userType,assignedLicenses&`$top=999"
if ($usersResult) {
    $users = @($usersResult.Items | Where-Object { $_.userPrincipalName -notlike "*#EXT#*" -and $_.userType -ne 'Guest' })
    $totalUsers = $users.Count
    Write-Log "Enabled member users: $totalUsers" "SUCCESS"
}
Set-Source "Users" ($null -ne $totalUsers)

# --- MFA registration -------------------------------------------------------
Write-Log "Collecting MFA registration..."
$mfaById = @{}
$mfaCollected = $false

$regResult = Get-GraphCollection -Token $Token -Endpoint "/reports/authenticationMethods/userRegistrationDetails?`$top=999"
if ($regResult) {
    foreach ($r in $regResult.Items) { $mfaById[[string]$r.id] = [bool]$r.isMfaRegistered }
    $mfaCollected = $true
    Write-Log "MFA registration details: $($regResult.Items.Count) records" "SUCCESS"
}
elseif ($users.Count -gt 0) {
    # Fallback (tenants without access to the registration report): per-user methods
    Write-Log "userRegistrationDetails unavailable - falling back to per-user authentication methods" "WARN"
    $okCalls = 0
    foreach ($u in $users) {
        $methods = Invoke-GraphRequest -Token $Token -Endpoint "/users/$($u.id)/authentication/methods" -Quiet
        if ($null -ne $methods) {
            $okCalls++
            $nonPassword = @($methods.value | Where-Object { $_.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod' })
            $mfaById[[string]$u.id] = ($nonPassword.Count -gt 0)
        }
    }
    $mfaCollected = ($okCalls -gt 0 -and $okCalls -eq $users.Count)
    if (-not $mfaCollected) { Write-Log "Per-user authentication methods incomplete ($okCalls of $($users.Count))" "WARN" }
}

$usersWithMfa = $null
$usersWithoutMfa = @()
$mfaPercent = $null
if ($mfaCollected -and $null -ne $totalUsers) {
    $usersWithMfa = 0
    foreach ($u in $users) {
        if ($mfaById[[string]$u.id]) { $usersWithMfa++ } else { $usersWithoutMfa += $u.displayName }
    }
    $mfaPercent = if ($totalUsers -gt 0) { [int][math]::Round(($usersWithMfa / $totalUsers) * 100, 0) } else { 0 }
    Write-Log "MFA: $usersWithMfa of $totalUsers ($mfaPercent%)" "SUCCESS"
}
Set-Source "MFA" ($null -ne $mfaPercent)

# --- Administrators (members of activated admin directory roles) ------------
Write-Log "Collecting administrators..."
$admins = @{}
$adminsCollected = $false
$rolesResult = Get-GraphCollection -Token $Token -Endpoint "/directoryRoles?`$select=id,displayName,roleTemplateId"
if ($rolesResult) {
    $adminsCollected = $true
    foreach ($role in @($rolesResult.Items | Where-Object { $_.displayName -match 'Administrator' })) {
        $members = Get-GraphCollection -Token $Token -Endpoint "/directoryRoles/$($role.id)/members?`$select=id,displayName,userPrincipalName"
        if ($null -eq $members) { $adminsCollected = $false; continue }
        foreach ($m in @($members.Items | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' })) {
            $admins[[string]$m.id] = $m
        }
    }
    Write-Log "Admin role members: $($admins.Count)" "SUCCESS"
}

$adminsTotal = $null
$adminsWithMfa = $null
$adminsWithoutMfa = @()
if ($adminsCollected -and $mfaCollected) {
    $adminsTotal = $admins.Count
    $adminsWithMfa = 0
    foreach ($a in $admins.Values) {
        if ($mfaById[[string]$a.id]) { $adminsWithMfa++ }
        else { $adminsWithoutMfa += $(if ($a.displayName) { $a.displayName } else { $a.userPrincipalName }) }
    }
    Write-Log "Admins with MFA: $adminsWithMfa of $adminsTotal" "SUCCESS"
}
Set-Source "AdminMFA" ($null -ne $adminsTotal)

# --- Sign-in logs (last 30 days, requires Entra ID P1/P2) -------------------
Write-Log "Collecting sign-in logs (last 30 days)..."
$signInSuccess = $null
$signInFailed = $null
$signInTruncated = $false
$signInCount = 0
$signInResult = Get-GraphCollection -Token $Token -Endpoint "/auditLogs/signIns?`$filter=createdDateTime ge $since30Iso&`$top=1000" -MaxPages $MaxSignInPages
if ($signInResult) {
    $signInCount = $signInResult.Items.Count
    $signInSuccess = @($signInResult.Items | Where-Object { $_.status.errorCode -eq 0 }).Count
    $signInFailed = $signInCount - $signInSuccess
    $signInTruncated = $signInResult.Truncated
    Write-Log "Sign-ins (30 days): $signInSuccess successful, $signInFailed failed/interrupted$(if ($signInTruncated) { ' (partial)' })" "SUCCESS"
}
else {
    Write-Log "Sign-in logs not available - requires Entra ID P1/P2 and AuditLog.Read.All" "WARN"
}
Set-Source "SignIns" ($null -ne $signInSuccess)

# --- Security alerts (alerts_v2, last 30 days) ------------------------------
Write-Log "Collecting security alerts..."
$alertsTotal = $null
$alertsHigh = $null
$alertsActive = $null
$alertsResult = Get-GraphCollection -Token $Token -Endpoint "/security/alerts_v2?`$filter=createdDateTime ge $since30Iso&`$top=1000"
if ($alertsResult) {
    $alertsTotal = $alertsResult.Items.Count
    $alertsHigh = @($alertsResult.Items | Where-Object { $_.severity -eq 'high' }).Count
    $alertsActive = @($alertsResult.Items | Where-Object { $_.status -ne 'resolved' }).Count
    Write-Log "Security alerts: $alertsTotal (high: $alertsHigh, unresolved: $alertsActive)" "SUCCESS"
}
Set-Source "SecurityAlerts" ($null -ne $alertsTotal)

# --- Service health ---------------------------------------------------------
Write-Log "Collecting service health..."
$healthServices = $null
$healthDegraded = $null
$incidents30 = $null
$advisories30 = $null
$healthResult = Get-GraphCollection -Token $Token -Endpoint "/admin/serviceAnnouncement/healthOverviews?`$select=service,status"
if ($healthResult) {
    $okStates = @('serviceOperational', 'serviceRestored', 'falsePositive', 'postIncidentReviewPublished', 'resolved')
    $healthServices = $healthResult.Items.Count
    $healthDegraded = @($healthResult.Items | Where-Object { $okStates -notcontains $_.status }).Count

    $issuesResult = Get-GraphCollection -Token $Token -Endpoint "/admin/serviceAnnouncement/issues?`$filter=startDateTime ge $since30Iso" -Quiet
    if ($null -eq $issuesResult) {
        $issuesResult = Get-GraphCollection -Token $Token -Endpoint "/admin/serviceAnnouncement/issues" -MaxPages 20
    }
    if ($issuesResult) {
        $recent = @($issuesResult.Items | Where-Object { $_.startDateTime -and ([datetime]$_.startDateTime).ToUniversalTime() -ge $since30 })
        $incidents30 = @($recent | Where-Object { $_.classification -eq 'incident' }).Count
        $advisories30 = @($recent | Where-Object { $_.classification -eq 'advisory' }).Count
    }
    Write-Log "Service health: $healthServices services, $healthDegraded with issues; incidents (30d): $incidents30" "SUCCESS"
}
Set-Source "ServiceHealth" ($null -ne $healthServices)

# --- Licenses ---------------------------------------------------------------
Write-Log "Collecting licenses..."
$licensesTotal = $null
$licensesUsed = $null
$licenseDetails = @()

# Free / internal SKUs ignored in the report
$skusToIgnore = @(
    "*FLOW_FREE*",              # Power Automate Free
    "*POWERAUTOMATE_FREE*",     # Power Automate Free (alternate name)
    "*TEAMS_EXPLORATORY*",      # Teams Exploratory
    "*POWERAPPS_VIRAL*",        # PowerApps Trial
    "*POWER_BI_STANDARD*",      # Power BI Free
    "*RIGHTSMANAGEMENT_ADHOC*", # RMS Ad-hoc
    "*WINDOWS_STORE*",          # Windows Store for Business
    "*STREAM*",                 # Stream Trial
    "*FORMS_PRO*",              # Forms Pro Trial
    "*CCIBOTS*",                # Power Virtual Agents Trial
    "*CLIPCHAMP*",              # Clipchamp
    "*Dynamics_365*TRIAL*",     # Dynamics trials
    "*VISIOCLIENT_FIRSTLINE*"   # Visio web free
)

$skuResult = Get-GraphCollection -Token $Token -Endpoint "/subscribedSkus"
if ($skuResult) {
    $licensesTotal = 0
    $licensesUsed = 0
    foreach ($sku in $skuResult.Items) {
        if ($sku.capabilityStatus -ne "Enabled") { continue }

        $ignore = $false
        foreach ($pattern in $skusToIgnore) {
            if ($sku.skuPartNumber -like $pattern) { $ignore = $true; break }
        }
        if ($ignore) {
            Write-Log "Ignoring free SKU: $($sku.skuPartNumber)"
            continue
        }

        $licensesTotal += [int]$sku.prepaidUnits.enabled
        $licensesUsed += [int]$sku.consumedUnits

        # First match wins (specific patterns before generic ones)
        $skuName = switch -Wildcard ($sku.skuPartNumber) {
            "*O365_BUSINESS_ESSENTIALS*" { "Office 365 Business Essentials"; break }
            "*O365_BUSINESS_PREMIUM*" { "Office 365 Business Premium"; break }
            "*BUSINESS_BASIC*" { "Microsoft 365 Business Basic"; break }
            "*BUSINESS_STANDARD*" { "Microsoft 365 Business Standard"; break }
            "*BUSINESS_PREMIUM*" { "Microsoft 365 Business Premium"; break }
            "*EXCHANGESTANDARD*" { "Exchange Online Plan 1"; break }
            "*EXCHANGEENTERPRISE*" { "Exchange Online Plan 2"; break }
            "*SHAREPOINTSTANDARD*" { "SharePoint Online Plan 1"; break }
            "*SHAREPOINTENTERPRISE*" { "SharePoint Online Plan 2"; break }
            "*AAD_PREMIUM*" { "Microsoft Entra ID Premium"; break }
            "*INTUNE*" { "Microsoft Intune"; break }
            "*DEFENDER*" { "Microsoft Defender"; break }
            "*ATP*" { "Microsoft Defender for Office 365"; break }
            "*EMS*" { "Enterprise Mobility + Security"; break }
            "*E1*" { "Microsoft 365 E1"; break }
            "*E3*" { "Microsoft 365 E3"; break }
            "*E5*" { "Microsoft 365 E5"; break }
            "*F1*" { "Microsoft 365 F1"; break }
            "*F3*" { "Microsoft 365 F3"; break }
            default { $sku.skuPartNumber }
        }

        $licenseDetails += [PSCustomObject]@{
            Name      = $skuName
            Total     = [int]$sku.prepaidUnits.enabled
            Assigned  = [int]$sku.consumedUnits
            Available = [int]$sku.prepaidUnits.enabled - [int]$sku.consumedUnits
        }
    }
    Write-Log "Licenses: $licensesUsed of $licensesTotal assigned" "SUCCESS"
}
Set-Source "Licenses" ($null -ne $licensesTotal)

# --- Usage reports (Reports.Read.All) ---------------------------------------
Write-Log "Collecting usage reports..."

$onedriveBytes = Get-LatestStorageBytes (Get-UsageReport -Token $Token -Function "getOneDriveUsageStorage(period='D7')")

$sharepointBytes = $null
$spRows = Get-UsageReport -Token $Token -Function "getSharePointSiteUsageDetail(period='D180')"
if ($null -ne $spRows) {
    $spActive = @($spRows | Where-Object { -not (Test-Truthy (Get-Field $_ @('Is Deleted', 'isDeleted'))) })
    $sharepointBytes = [double]0
    foreach ($row in $spActive) {
        $b = ConvertTo-Double (Get-Field $row @('Storage Used (Byte)', 'storageUsedInBytes'))
        if ($null -ne $b) { $sharepointBytes += $b }
    }
    Write-Log "SharePoint: $($spActive.Count) sites"
}

$exchangeBytes = Get-LatestStorageBytes (Get-UsageReport -Token $Token -Function "getMailboxUsageStorage(period='D7')")

$totalMailboxes = $null
$sharedMailboxes = $null
$mbRows = Get-UsageReport -Token $Token -Function "getMailboxUsageDetail(period='D30')"
if ($null -ne $mbRows) {
    $mbActive = @($mbRows | Where-Object { -not (Test-Truthy (Get-Field $_ @('Is Deleted', 'isDeleted'))) })
    $totalMailboxes = $mbActive.Count
    # Shared mailbox count only when the report exposes the recipient type.
    $withType = @($mbActive | Where-Object { $null -ne (Get-Field $_ @('Recipient Type', 'recipientType')) })
    if ($mbActive.Count -gt 0 -and $withType.Count -eq $mbActive.Count) {
        $sharedMailboxes = @($withType | Where-Object { (Get-Field $_ @('Recipient Type', 'recipientType')) -match 'Shared' }).Count
    }
}

$serviceUsage = @()
$svcRows = Get-UsageReport -Token $Token -Function "getOffice365ServicesUserCounts(period='D30')"
if ($null -ne $svcRows -and @($svcRows).Count -gt 0) {
    $row = @($svcRows)[0]
    $svcDefs = @(
        @{ Name = "Exchange"; Active = @('Exchange Active', 'exchangeActive'); Inactive = @('Exchange Inactive', 'exchangeInactive') },
        @{ Name = "OneDrive"; Active = @('OneDrive Active', 'oneDriveActive'); Inactive = @('OneDrive Inactive', 'oneDriveInactive') },
        @{ Name = "SharePoint"; Active = @('SharePoint Active', 'sharePointActive'); Inactive = @('SharePoint Inactive', 'sharePointInactive') },
        @{ Name = "Teams"; Active = @('Teams Active', 'teamsActive'); Inactive = @('Teams Inactive', 'teamsInactive') },
        @{ Name = "Viva Engage (Yammer)"; Active = @('Yammer Active', 'yammerActive'); Inactive = @('Yammer Inactive', 'yammerInactive') }
    )
    foreach ($svc in $svcDefs) {
        $active = ConvertTo-Double (Get-Field $row $svc.Active)
        if ($null -eq $active) { continue }
        $inactive = ConvertTo-Double (Get-Field $row $svc.Inactive)
        if ($null -ne $inactive -and ($active + $inactive) -eq 0) { continue }   # service not in use / not licensed
        $denominator = if ($null -ne $inactive -and ($active + $inactive) -gt 0) { $active + $inactive } elseif ($totalUsers) { $totalUsers } else { 0 }
        if ($active -eq 0 -and $denominator -eq 0) { continue }
        $pct = if ($denominator -gt 0) { [int][math]::Round(($active / $denominator) * 100, 0) } else { $null }
        $serviceUsage += [PSCustomObject]@{ Service = $svc.Name; Active = $active; Percent = $pct }
    }
}

Set-Source "Storage" ($null -ne $onedriveBytes -or $null -ne $sharepointBytes -or $null -ne $exchangeBytes)
Set-Source "ServiceUsage" ($serviceUsage.Count -gt 0)

# --- Groups and Teams -------------------------------------------------------
Write-Log "Collecting groups and Teams..."
$totalGroups = $null
$totalTeams = $null
$groupsResult = Get-GraphCollection -Token $Token -Endpoint "/groups?`$filter=groupTypes/any(c:c eq 'Unified')&`$select=id,displayName,resourceProvisioningOptions&`$top=999"
if ($groupsResult) {
    $totalGroups = $groupsResult.Items.Count
    $totalTeams = @($groupsResult.Items | Where-Object { $_.resourceProvisioningOptions -contains "Team" }).Count
    Write-Log "Microsoft 365 groups: $totalGroups, Teams: $totalTeams" "SUCCESS"
}
Set-Source "Groups" ($null -ne $totalGroups)

Write-Log "Data collection finished" "SUCCESS"

# ============================================================================
# BUILD HTML
# ============================================================================

Write-Log "Building HTML report..."

$templateHtml = Get-Content $TemplatePath -Raw -Encoding UTF8
$NC = T 'NotCollected'

function Show-Value {
    param($Value, [string]$Format = 'num')
    if ($null -eq $Value) { return $NC }
    switch ($Format) {
        'num' { return (Format-Num $Value) }
        'pct' { return "$Value%" }
        default { return [string]$Value }
    }
}

function Get-ScoreClass {
    param($Score)
    if ($null -eq $Score) { return "neutral" }
    if ($Score -ge 70) { return "success" }
    elseif ($Score -ge 40) { return "warning" }
    else { return "danger" }
}

function Get-BarClass {
    param($Score)
    if ($null -eq $Score) { return "low" }
    if ($Score -ge 70) { return "" }
    elseif ($Score -ge 40) { return "medium" }
    else { return "low" }
}

# --- Logo (msp.logo in config; falls back to templates\logo.png if present) --
$logoImg = ""
$logoCandidates = @()
if ($MspLogo) {
    if ([System.IO.Path]::IsPathRooted([string]$MspLogo)) { $logoCandidates += [string]$MspLogo }
    else { $logoCandidates += (Join-Path $ConfigDir $MspLogo); $logoCandidates += (Join-Path $ScriptPath $MspLogo) }
}
else {
    $logoCandidates += (Join-Path (Join-Path $ScriptPath "templates") "logo.png")
}
foreach ($lp in $logoCandidates) {
    if (Test-Path $lp) {
        $ext = [System.IO.Path]::GetExtension($lp).ToLowerInvariant()
        $mime = switch ($ext) { ".svg" { "image/svg+xml" } ".jpg" { "image/jpeg" } ".jpeg" { "image/jpeg" } ".gif" { "image/gif" } default { "image/png" } }
        $b64 = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes((Resolve-Path $lp).Path))
        $logoImg = "<img src=`"data:$mime;base64,$b64`" alt=`"Logo`" class=`"header-logo`">"
        Write-Log "Logo loaded: $lp" "SUCCESS"
        break
    }
}

$V = @{}   # placeholder -> already-safe HTML

$V.HTML_LANG = Enc $(if ($Language -eq 'en') { 'en' } else { $Language })
$V.LOGO_IMG = $logoImg
$V.MSP_HEADER_BLOCK = if ($MspName) { "<div class=`"logo-text`">$(Enc $MspName)</div>" } else { "" }
$V.CLIENT_NAME = Enc $ClientName
$V.PERIOD = Enc $period

# Executive summary
$okSources = @($Sources.Values | Where-Object { $_ }).Count
$allSources = $Sources.Count
if ($okSources -eq $allSources) {
    $V.EXEC_SOURCES_CLASS = "green"; $V.EXEC_SOURCES_ICON = "&#10003;"
    $V.EXEC_SOURCES_TEXT = Enc (T 'ExecSourcesAll' $allSources)
}
else {
    $V.EXEC_SOURCES_CLASS = "yellow"; $V.EXEC_SOURCES_ICON = "!"
    $V.EXEC_SOURCES_TEXT = Enc (T 'ExecSourcesPartial' $okSources, $allSources)
}
$V.EXEC_USERS_TEXT = if ($null -ne $totalUsers) { Enc (T 'ExecUsers' (Format-Num $totalUsers)) } else { Enc (T 'ExecUsersUnknown') }

if ($null -eq $secureScore) {
    $V.EXEC_SECURITY_CLASS = "gray"; $V.EXEC_SECURITY_ICON = "?"; $V.EXEC_SECURITY_TEXT = Enc (T 'SecurityUnknown')
}
elseif ($secureScore -ge 50 -and $adminsWithoutMfa.Count -eq 0) {
    $V.EXEC_SECURITY_CLASS = "green"; $V.EXEC_SECURITY_ICON = "&#10003;"; $V.EXEC_SECURITY_TEXT = Enc (T 'SecurityAdequate')
}
else {
    $V.EXEC_SECURITY_CLASS = "yellow"; $V.EXEC_SECURITY_ICON = "!"; $V.EXEC_SECURITY_TEXT = Enc (T 'SecurityAttention')
}

# Service health
if ($null -eq $healthServices) {
    $V.HEALTH_CLASS = "unknown"
    $V.HEALTH_DETAIL = Enc (T 'HealthNotCollected')
    $V.HEALTH_INCIDENTS = Enc $NC
}
else {
    $V.HEALTH_CLASS = if ($healthDegraded -gt 0) { "warn" } else { "" }
    $V.HEALTH_DETAIL = if ($healthDegraded -gt 0) { Enc (T 'HealthIssues' $healthDegraded, $healthServices) } else { Enc (T 'HealthAllOk' $healthServices) }
    $V.HEALTH_INCIDENTS = if ($null -ne $incidents30) { Enc (T 'HealthIncidentsFmt' $incidents30, $advisories30) } else { Enc $NC }
}

# Secure Score
$V.SECURE_SCORE_DISPLAY = if ($null -ne $secureScore) { "$secureScore%" } else { Enc (T 'NotCollected') }
$V.SECURE_SCORE_CLASS = Get-ScoreClass $secureScore
$V.SECURE_SCORE_BAR_CLASS = Get-BarClass $secureScore
$V.SECURE_SCORE_WIDTH = if ($null -ne $secureScore) { [string]([math]::Max([math]::Min($secureScore, 100), 0)) } else { "0" }
$V.SECURE_SCORE_CURRENT = if ($null -ne $secureScoreCurrent) { Enc (T 'Points' $secureScoreCurrent) } else { Enc $NC }
$V.SECURE_SCORE_MAX = if ($null -ne $secureScoreMax) { Enc (T 'Points' $secureScoreMax) } else { Enc $NC }
if ($null -ne $secureScoreChange) {
    $V.SECURE_SCORE_CHANGE = if ($secureScoreChange -ge 0) { "+$secureScoreChange%" } else { "$secureScoreChange%" }
    $V.SECURE_SCORE_CHANGE_CLASS = if ($secureScoreChange -gt 0) { "good" } elseif ($secureScoreChange -lt 0) { "bad" } else { "" }
}
else {
    $V.SECURE_SCORE_CHANGE = Enc $(if ($null -ne $secureScore) { "N/A" } else { $NC })
    $V.SECURE_SCORE_CHANGE_CLASS = "muted"
}

# Users / MFA
$V.TOTAL_USERS = Enc (Show-Value $totalUsers)
$V.USERS_CLASS = if ($null -ne $totalUsers) { "info" } else { "neutral" }
$V.MFA_DISPLAY = if ($null -ne $mfaPercent) { "$mfaPercent%" } else { Enc $NC }
$V.MFA_CLASS = Get-ScoreClass $mfaPercent
$V.MFA_VALUE_CLASS = if ($null -eq $mfaPercent) { "muted" } elseif ($mfaPercent -ge 80) { "good" } elseif ($mfaPercent -ge 50) { "warning" } else { "bad" }
$V.USERS_MFA = if ($null -ne $usersWithMfa) { Enc (T 'OfFormat' (Format-Num $usersWithMfa), (Format-Num $totalUsers)) } else { Enc $NC }

if ($null -eq $adminsTotal) {
    $V.ADMINS_MFA = Enc $NC; $V.ADMINS_MFA_CLASS = "muted"
}
elseif ($adminsTotal -eq 0) {
    $V.ADMINS_MFA = Enc (T 'NoAdminsFound'); $V.ADMINS_MFA_CLASS = ""
}
else {
    $V.ADMINS_MFA = Enc (T 'OfFormat' $adminsWithMfa, $adminsTotal)
    $V.ADMINS_MFA_CLASS = if ($adminsWithMfa -eq $adminsTotal) { "good" } else { "bad" }
}

$mfaAlerts = ""
if ($adminsWithoutMfa.Count -gt 0) {
    $mfaAlerts += "<div class=`"alert-box`"><h4>$(Enc (T 'AdminsWithoutMfa' $adminsWithoutMfa.Count))</h4><ul>`n"
    foreach ($n in ($adminsWithoutMfa | Sort-Object)) { $mfaAlerts += "<li>$(Enc $n)</li>`n" }
    $mfaAlerts += "</ul></div>`n"
}
if ($IncludeUsersWithoutMfa -and $usersWithoutMfa.Count -gt 0) {
    $mfaAlerts += "<div class=`"alert-box`"><h4>$(Enc (T 'UsersWithoutMfa' $usersWithoutMfa.Count))</h4><ul>`n"
    foreach ($n in ($usersWithoutMfa | Sort-Object)) { $mfaAlerts += "<li>$(Enc $n)</li>`n" }
    $mfaAlerts += "</ul></div>`n"
}
$V.MFA_ALERTS_SECTION = $mfaAlerts

# Sign-ins
$plus = if ($signInTruncated) { "+" } else { "" }
if ($null -eq $signInSuccess) {
    $V.SIGNINS_TOTAL = Enc $NC; $V.SIGNINS_SUCCESS = Enc $NC; $V.SIGNINS_FAILED = Enc $NC
    $V.SIGNINS_FAILED_CLASS = "neutral"; $V.SIGNINS_FAILED_VALUE_CLASS = "muted"
    $V.SIGNINS_NOTE = "<p class=`"note`">$(Enc (T 'SigninsNotCollected'))</p>"
}
else {
    $V.SIGNINS_TOTAL = Enc ((Format-Num $signInCount) + $plus)
    $V.SIGNINS_SUCCESS = Enc ((Format-Num $signInSuccess) + $plus)
    $V.SIGNINS_FAILED = Enc ((Format-Num $signInFailed) + $plus)
    if ($signInFailed -gt 100) { $V.SIGNINS_FAILED_CLASS = "danger"; $V.SIGNINS_FAILED_VALUE_CLASS = "bad" }
    elseif ($signInFailed -gt 50) { $V.SIGNINS_FAILED_CLASS = "warning"; $V.SIGNINS_FAILED_VALUE_CLASS = "warning" }
    else { $V.SIGNINS_FAILED_CLASS = "success"; $V.SIGNINS_FAILED_VALUE_CLASS = "" }
    $note = T 'SigninsNote'
    if ($signInTruncated) { $note = (T 'SigninsTruncated' (Format-Num $signInCount)) + " " + $note }
    $V.SIGNINS_NOTE = "<p class=`"note`">$(Enc $note)</p>"
}

# Security alerts
if ($null -eq $alertsTotal) {
    $V.ALERTS_TOTAL = "<span class=`"text-muted`">&#8211;</span>"
    $V.ALERTS_HIGH = "<span class=`"text-muted`">&#8211;</span>"
    $V.ALERTS_ACTIVE = "<span class=`"text-muted`">&#8211;</span>"
    $V.SECURITY_STATUS = Enc (T 'AlertsNotCollected'); $V.SECURITY_STATUS_CLASS = "muted"
}
else {
    $V.ALERTS_TOTAL = "<span class=`"$(if ($alertsTotal -gt 0) { 'text-yellow' } else { 'text-green' })`">$(Format-Num $alertsTotal)</span>"
    $V.ALERTS_HIGH = "<span class=`"$(if ($alertsHigh -gt 0) { 'text-red' } else { 'text-green' })`">$(Format-Num $alertsHigh)</span>"
    $V.ALERTS_ACTIVE = "<span class=`"$(if ($alertsActive -gt 0) { 'text-red' } else { 'text-green' })`">$(Format-Num $alertsActive)</span>"
    if ($alertsTotal -eq 0) { $V.SECURITY_STATUS = Enc (T 'AlertsNone'); $V.SECURITY_STATUS_CLASS = "good" }
    elseif ($alertsActive -gt 0) { $V.SECURITY_STATUS = Enc (T 'AlertsReview' $alertsActive); $V.SECURITY_STATUS_CLASS = "bad" }
    else { $V.SECURITY_STATUS = Enc (T 'AlertsAllResolved' $alertsTotal); $V.SECURITY_STATUS_CLASS = "good" }
}

# Collaboration
$V.TOTAL_TEAMS = Enc (Show-Value $totalTeams)
$V.TOTAL_GROUPS = Enc (Show-Value $totalGroups)
$V.TOTAL_MAILBOXES = Enc (Show-Value $totalMailboxes)
$V.SHARED_MAILBOXES = Enc (Show-Value $sharedMailboxes)

# Licenses
$V.LICENSES_TOTAL = Enc (Show-Value $licensesTotal)
$V.LICENSES_USED = Enc (Show-Value $licensesUsed)
$V.LICENSES_AVAILABLE = if ($null -ne $licensesTotal) { Enc (Format-Num ($licensesTotal - $licensesUsed)) } else { Enc $NC }
$V.LICENSES_UTILIZATION = if ($licensesTotal -gt 0) { "$([int][math]::Round(($licensesUsed / $licensesTotal) * 100, 0))%" } elseif ($null -ne $licensesTotal) { "0%" } else { Enc $NC }

$licHtml = ""
if ($licenseDetails.Count -gt 0) {
    $licHtml = "<table style='margin-top: 15px;'><thead><tr><th>$(Enc (T 'LicPlan'))</th><th>$(Enc (T 'LicTotal'))</th><th>$(Enc (T 'LicAssigned'))</th><th>$(Enc (T 'LicAvailable'))</th></tr></thead><tbody>"
    foreach ($lic in $licenseDetails) {
        $licHtml += "<tr><td>$(Enc $lic.Name)</td><td>$(Format-Num $lic.Total)</td><td>$(Format-Num $lic.Assigned)</td><td>$(Format-Num $lic.Available)</td></tr>"
    }
    $licHtml += "</tbody></table>"
}
$V.LICENSES_DETAILS = $licHtml

# Storage
$V.ONEDRIVE_USED = if ($null -ne $onedriveBytes) { Enc (Format-Bytes $onedriveBytes) } else { Enc $NC }
$V.SHAREPOINT_USED = if ($null -ne $sharepointBytes) { Enc (Format-Bytes $sharepointBytes) } else { Enc $NC }
$V.EXCHANGE_USED = if ($null -ne $exchangeBytes) { Enc (Format-Bytes $exchangeBytes) } else { Enc $NC }
$storageParts = @($onedriveBytes, $sharepointBytes, $exchangeBytes | Where-Object { $null -ne $_ })
if ($storageParts.Count -eq 0) { $V.STORAGE_TOTAL = Enc $NC }
else {
    $total = Format-Bytes (($storageParts | Measure-Object -Sum).Sum)
    $V.STORAGE_TOTAL = if ($storageParts.Count -lt 3) { Enc (T 'Partial' $total) } else { Enc $total }
}

# Service usage
$rows = ""
foreach ($svc in $serviceUsage) {
    $pctText = if ($null -ne $svc.Percent) { "$($svc.Percent)%" } else { "N/A" }
    $rows += "<tr><td>$(Enc $svc.Service)</td><td>$(Format-Num $svc.Active)</td><td>$pctText</td></tr>`n"
}
if ($rows -eq "") {
    $rows = "<tr><td colspan='3' style='text-align:center; color:#A0A0A0;'>$(Enc (T 'UsageNotAvailable'))</td></tr>"
}
$V.SERVICE_USAGE_ROWS = $rows

# Recommendations (derived only from collected data)
$recommendations = @()
if ($adminsWithoutMfa.Count -gt 0) { $recommendations += T 'RecAdminMfa' $adminsWithoutMfa.Count }
if ($null -ne $mfaPercent -and $mfaPercent -lt 100) { $recommendations += T 'RecMfa' $mfaPercent }
if ($null -ne $secureScore -and $secureScore -lt 50) { $recommendations += T 'RecSecureScore' }
if ($null -ne $alertsActive -and $alertsActive -gt 0) { $recommendations += T 'RecAlerts' $alertsActive }
if ($null -ne $signInFailed -and $signInFailed -gt 50) { $recommendations += T 'RecFailedSignins' (Format-Num $signInFailed) }
if ($null -eq $signInSuccess) { $recommendations += T 'RecSigninsP1' }

$recSection = ""
if ($IncludeRecommendations -and $recommendations.Count -gt 0) {
    $recSection = "<div class=`"section`">`n    <h2 class=`"section-title`">$(Enc (T 'RecTitle'))</h2>`n    <div class=`"recommendations-box`">`n        <h4>$(Enc (T 'RecSubtitle'))</h4>`n        <ul>`n"
    foreach ($rec in $recommendations) { $recSection += "<li>$(Enc $rec)</li>`n" }
    $recSection += "</ul></div></div>"
}
$V.RECOMMENDATIONS_SECTION = $recSection

# Technical summary (derived only from collected data)
$summary = @()
if ($null -eq $secureScore) { $summary += T 'SumScoreNA' }
elseif ($secureScore -ge 70) { $summary += T 'SumScoreGood' $secureScore }
elseif ($secureScore -ge 40) { $summary += T 'SumScoreMid' $secureScore }
else { $summary += T 'SumScoreLow' $secureScore }

if ($null -ne $mfaPercent) { $summary += T 'SumMfa' $mfaPercent } else { $summary += T 'SumMfaNA' }
if ($null -ne $adminsTotal -and $adminsTotal -gt 0) {
    if ($adminsWithoutMfa.Count -eq 0) { $summary += T 'SumAdminsOk' $adminsTotal }
    else { $summary += T 'SumAdminsBad' $adminsWithoutMfa.Count, $adminsTotal }
}
if ($null -eq $alertsTotal) { $summary += T 'SumAlertsNA' }
elseif ($alertsTotal -eq 0) { $summary += T 'SumAlertsNone' }
else { $summary += T 'SumAlerts' $alertsTotal, $alertsHigh, $alertsActive }
if ($null -ne $incidents30 -and $incidents30 -gt 0) { $summary += T 'SumIncidents' $incidents30 }
if ($null -ne $signInFailed) { $summary += T 'SumSignins' ((Format-Num $signInFailed) + $plus) }
$V.SUMMARY_TEXT = Enc ($summary -join " ")

# MSP branding / footer
$V.MSP_NAME = Enc $MspName
$V.MSP_TAGLINE_BLOCK = if ($MspTagline) { "<div class=`"footer-tagline`">$(Enc $MspTagline)</div>" } else { "" }
$V.MSP_CONTACT = Enc $MspContact
$V.GENERATED_AT = Enc ((Get-Date).ToString((T 'DateFormat'), $Culture))

# Single-pass substitution: values are never re-scanned for placeholders,
# and no regex replacement strings are used (so '$' in data is safe).
$html = [regex]::Replace($templateHtml, '\{\{(L:)?([A-Z0-9_]+)\}\}', {
        param($m)
        if ($m.Groups[1].Success) {
            $s = $Strings[$m.Groups[2].Value]
            if ($null -ne $s) { return [System.Net.WebUtility]::HtmlEncode($s) }
            return $m.Value
        }
        $key = $m.Groups[2].Value
        if ($V.ContainsKey($key)) { return [string]$V[$key] }
        return $m.Value
    })

$leftover = [regex]::Matches($html, '\{\{[A-Z0-9_:]+\}\}') | ForEach-Object { $_.Value } | Select-Object -Unique
if ($leftover) { Write-Log "Unreplaced template placeholders: $($leftover -join ', ')" "WARN" }

# ============================================================================
# SAVE HTML AND GENERATE PDF
# ============================================================================

$timestamp = Get-Date -Format "yyyyMM"
$safeName = $ClientName.Replace(' ', '_')
foreach ($c in ([System.IO.Path]::GetInvalidFileNameChars() + '<>:"/\|?*'.ToCharArray())) { $safeName = $safeName.Replace([string]$c, '_') }
if (-not $safeName) { $safeName = "Client" }
$fileBase = "${safeName}_M365_Report_$timestamp"
$htmlFile = Join-Path $OutputPath "$fileBase.html"
$pdfFile = Join-Path $OutputPath "$fileBase.pdf"

$html | Out-File -FilePath $htmlFile -Encoding UTF8
Write-Log "HTML saved: $htmlFile" "SUCCESS"

if ($Preview) {
    Write-Log "Opening preview in the browser..."
    Start-Process $htmlFile
    Write-Log "Preview opened. PDF not generated (preview mode)." "SUCCESS"
    exit 0
}

Write-Log "Generating PDF with Microsoft Edge..."

$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (!(Test-Path $edgePath)) {
    $edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
}

if (Test-Path $edgePath) {
    $htmlFullPath = (Resolve-Path $htmlFile).Path
    $pdfFullPath = Join-Path (Resolve-Path $OutputPath).Path "$fileBase.pdf"

    $arguments = "--headless --disable-gpu --print-to-pdf=`"$pdfFullPath`" --no-pdf-header-footer `"file:///$($htmlFullPath.Replace('\', '/'))`""
    Start-Process -FilePath $edgePath -ArgumentList $arguments -Wait -NoNewWindow
    Start-Sleep -Seconds 2

    if (Test-Path $pdfFullPath) {
        Write-Log "PDF generated: $pdfFullPath" "SUCCESS"
        $pdfFile = $pdfFullPath
    }
    else {
        Write-Log "Failed to generate PDF. Check Microsoft Edge." "ERROR"
        exit 1
    }
}
else {
    Write-Log "Microsoft Edge not found. Install Edge or use -Preview." "ERROR"
    exit 1
}

# ============================================================================
# SEND EMAIL (OPTIONAL)
# ============================================================================

if ($SendEmail) {
    Write-Log "Preparing email..."

    $recipients = @(Get-ConfigValue $Config @('email.recipients', 'email.destinatarios') @())
    $ccList = @(Get-ConfigValue $Config @('email.cc') @())
    $subjectTemplate = [string](Get-ConfigValue $Config @('email.subject_template', 'email.assunto_template') (T 'EmailSubject'))
    $subject = $subjectTemplate.Replace("{{PERIOD}}", $period).Replace("{{PERIODO}}", $period)

    $emailSent = $false
    $senderToken = $null

    if (-not $MspSender) {
        Write-Log "No sender configured (msp.sender). Email not sent." "ERROR"
    }
    elseif ($recipients.Count -eq 0) {
        Write-Log "No recipients configured (email.recipients). Email not sent." "ERROR"
    }
    else {
        # Sender app (MSP tenant). Required when the sender mailbox is not in the client tenant.
        $sTenant = Get-ConfigValue $Config @('msp.sender_app.tenant_id')
        $sClient = Get-ConfigValue $Config @('msp.sender_app.client_id')
        $sSecret = Get-ConfigValue $Config @('msp.sender_app.client_secret')
        $sThumb = Get-ConfigValue $Config @('msp.sender_app.certificate_thumbprint')

        if ($sTenant -and $sClient -and ($sSecret -or $sThumb)) {
            Write-Log "Authenticating sender app (tenant $sTenant)..."
            try { $senderToken = Get-GraphToken -TenantId $sTenant -ClientId $sClient -ClientSecret $sSecret -CertificateThumbprint $sThumb }
            catch { Write-Log "Sender app authentication failed. Check msp.sender_app in the config." "ERROR" }
        }
        else {
            # No sender app: only valid when the sender mailbox lives in the client tenant.
            $senderDomain = $MspSender.Split('@')[-1]
            $domains = Get-GraphCollection -Token $Token -Endpoint "/domains?`$select=id"
            if ($domains -and (@($domains.Items | ForEach-Object { $_.id }) -contains $senderDomain)) {
                Write-Log "Sender domain $senderDomain belongs to the client tenant - using the client app to send."
                $senderToken = $Token
            }
            elseif ($domains) {
                Write-Log ("Sender $MspSender is not in the client tenant, so the client app cannot send as it. " +
                    "Configure msp.sender_app (tenant_id, client_id, client_secret or certificate_thumbprint) " +
                    "for an app in the MSP tenant with Mail.Send.") "ERROR"
            }
            else {
                Write-Log "Could not verify the sender domain (Domain list unavailable). Trying the client app." "WARN"
                $senderToken = $Token
            }
        }
    }

    if ($senderToken) {
        $pdfBase64 = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($pdfFile))

        $summaryItems = "<li>Secure Score: $(if ($null -ne $secureScore) { "$secureScore%" } else { Enc $NC })</li>"
        $summaryItems += "<li>$(Enc (T 'EmailUsers')): $(Enc (Show-Value $totalUsers))</li>"
        $summaryItems += "<li>$(Enc (T 'EmailMfa')): $(if ($null -ne $mfaPercent) { "$mfaPercent%" } else { Enc $NC })</li>"
        $signature = "<strong>$(Enc $MspName)</strong>"
        if ($MspContact) { $signature += "<br>$(Enc $MspContact)" }

        $emailBody = @"
<html>
<body style="font-family: Segoe UI, sans-serif; color: #333;">
<p>$(Enc (T 'EmailGreeting'))</p>
<p>$(T 'EmailIntro' (Enc $period))</p>
<p><strong>$(Enc (T 'EmailSummary'))</strong></p>
<ul>
$summaryItems
</ul>
<p>$(Enc (T 'EmailQuestions'))</p>
<p>$(Enc (T 'EmailRegards'))<br>
$signature</p>
</body>
</html>
"@

        $message = @{
            message         = @{
                subject      = $subject
                body         = @{ contentType = "HTML"; content = $emailBody }
                toRecipients = @($recipients | ForEach-Object { @{ emailAddress = @{ address = [string]$_ } } })
                attachments  = @(
                    @{
                        "@odata.type" = "#microsoft.graph.fileAttachment"
                        name          = "$fileBase.pdf"
                        contentType   = "application/pdf"
                        contentBytes  = $pdfBase64
                    }
                )
            }
            saveToSentItems = $true
        }
        if ($ccList.Count -gt 0) {
            $message.message.ccRecipients = @($ccList | ForEach-Object { @{ emailAddress = @{ address = [string]$_ } } })
        }

        $sendResult = Invoke-GraphRequest -Token $senderToken -Endpoint "/users/$([uri]::EscapeDataString($MspSender).Replace('%40', '@'))/sendMail" -Method Post -Body $message
        if ($script:LastGraphStatus -eq 200) {
            $emailSent = $true
            Write-Log "Email sent to: $($recipients -join ', ')" "SUCCESS"
        }
        elseif ($script:LastGraphStatus -eq 403) {
            Write-Log ("sendMail returned 403 for $MspSender. The sending app needs the Mail.Send application " +
                "permission (admin consent) in the tenant that hosts the mailbox, and any Exchange application " +
                "access policy must include this mailbox.") "ERROR"
        }
        elseif ($script:LastGraphStatus -eq 404) {
            Write-Log "sendMail returned 404: mailbox $MspSender not found in the sending tenant." "ERROR"
        }
        else {
            Write-Log "Failed to send email (status $($script:LastGraphStatus))." "ERROR"
        }
        $null = $sendResult
    }

    if (-not $emailSent) { Write-Log "The PDF was generated at: $pdfFile" "WARN" }
}

# ============================================================================
# DONE
# ============================================================================

Write-Log "========================================"
Write-Log "REPORT COMPLETE"
Write-Log "Client: $ClientName"
Write-Log "Period: $period"
Write-Log "PDF: $pdfFile"
Write-Log "Log: $LogFile"
Write-Log "========================================"

return $pdfFile
