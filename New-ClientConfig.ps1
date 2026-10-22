<#
.SYNOPSIS
    Creates the configuration file for a new client.

.DESCRIPTION
    Writes config\<ClientName>.json. Certificate authentication is recommended
    (-CertificateThumbprint): no secret is stored in the file. When a client
    secret is used it is stored in plain text, so the file ACL is restricted
    to the current user and SYSTEM (Windows).

.PARAMETER ClientName
    Client name (used for the file name and in the report).

.PARAMETER TenantId
    Client tenant ID (Entra ID).

.PARAMETER ClientId
    Application (client) ID of the app registration in the client tenant.

.PARAMETER CertificateThumbprint
    Thumbprint of a certificate (CurrentUser\My or LocalMachine\My) uploaded to
    the app registration. Recommended.

.PARAMETER ClientSecret
    Client secret of the app registration. If neither this nor
    -CertificateThumbprint is given, the secret is prompted for.

.PARAMETER Email
    Report recipient(s).

.PARAMETER MSPName / MSPContact / MSPSender / MSPTagline
    MSP branding and sender mailbox. Prompted for when not provided
    (tagline is optional).

.PARAMETER SenderTenantId / SenderClientId / SenderCertificateThumbprint / SenderClientSecret
    Optional app registration in the MSP tenant used to send the email
    (needs Mail.Send). Required when the sender mailbox is not in the client tenant.

.PARAMETER Language
    Report language: "en" (default) or "pt-BR".

.EXAMPLE
    .\New-ClientConfig.ps1 -ClientName "Contoso" -TenantId "<tenant-id>" -ClientId "<client-id>" -CertificateThumbprint "<thumbprint>" -Email "it@contoso.com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientName,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $true)]
    [string[]]$Email,

    [Parameter(Mandatory = $false)]
    [string]$MSPName,

    [Parameter(Mandatory = $false)]
    [string]$MSPContact,

    [Parameter(Mandatory = $false)]
    [string]$MSPSender,

    [Parameter(Mandatory = $false)]
    [string]$MSPTagline = "",

    [Parameter(Mandatory = $false)]
    [string]$SenderTenantId,

    [Parameter(Mandatory = $false)]
    [string]$SenderClientId,

    [Parameter(Mandatory = $false)]
    [string]$SenderCertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$SenderClientSecret,

    [Parameter(Mandatory = $false)]
    [ValidateSet("en", "pt-BR")]
    [string]$Language = "en"
)

function Read-Required {
    param([string]$Prompt)
    do { $v = Read-Host $Prompt } while ([string]::IsNullOrWhiteSpace($v))
    return $v.Trim()
}

function Read-SecretText {
    param([string]$Prompt)
    $secure = Read-Host $Prompt -AsSecureString
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

if (-not $MSPName) { $MSPName = Read-Required "MSP name (shown in the report)" }
if (-not $MSPContact) { $MSPContact = Read-Required "MSP contact (shown in the report footer, e.g. support@yourmsp.com)" }
if (-not $MSPSender) { $MSPSender = Read-Required "Sender mailbox for the report email (e.g. reports@yourmsp.com)" }

if (-not $CertificateThumbprint -and -not $ClientSecret) {
    Write-Host "No -CertificateThumbprint given. Certificate auth is recommended; press Ctrl+C to cancel," -ForegroundColor Yellow
    Write-Host "or enter the client secret to continue with secret auth." -ForegroundColor Yellow
    $ClientSecret = Read-SecretText "Client secret"
}

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptPath "config"

if (!(Test-Path $ConfigPath)) {
    New-Item -ItemType Directory -Path $ConfigPath -Force | Out-Null
}

$safeFileName = $ClientName -replace '[^a-zA-Z0-9]', '_'
$configFile = Join-Path $ConfigPath "$safeFileName.json"

$client = [ordered]@{
    name      = $ClientName
    tenant_id = $TenantId
    client_id = $ClientId
}
if ($CertificateThumbprint) { $client.certificate_thumbprint = $CertificateThumbprint }
else { $client.client_secret = $ClientSecret }

$msp = [ordered]@{
    name    = $MSPName
    contact = $MSPContact
    sender  = $MSPSender
    tagline = $MSPTagline
    logo    = ""
}
if ($SenderTenantId -and $SenderClientId -and ($SenderCertificateThumbprint -or $SenderClientSecret)) {
    $senderApp = [ordered]@{ tenant_id = $SenderTenantId; client_id = $SenderClientId }
    if ($SenderCertificateThumbprint) { $senderApp.certificate_thumbprint = $SenderCertificateThumbprint }
    else { $senderApp.client_secret = $SenderClientSecret }
    $msp.sender_app = $senderApp
}

$config = [ordered]@{
    client  = $client
    msp     = $msp
    email   = [ordered]@{
        recipients       = @($Email)
        cc               = @()
        subject_template = $(if ($Language -eq "pt-BR") { "Relatorio Mensal Microsoft 365 - {{PERIOD}}" } else { "Microsoft 365 Monthly Report - {{PERIOD}}" })
    }
    options = [ordered]@{
        include_users_without_mfa = $true
        include_recommendations   = $true
        language                  = $Language
    }
}

$config | ConvertTo-Json -Depth 5 | Out-File -FilePath $configFile -Encoding UTF8

$hasSecret = (-not $CertificateThumbprint) -or ($msp.Contains('sender_app') -and -not $SenderCertificateThumbprint)
if ($hasSecret) {
    Write-Host ""
    Write-Host "[WARN] A client secret is stored in plain text in $configFile." -ForegroundColor Yellow
    Write-Host "       Never commit it. Prefer -CertificateThumbprint." -ForegroundColor Yellow
    if ($env:OS -eq "Windows_NT") {
        try {
            # Restrict the file to the current user and SYSTEM
            $acl = Get-Acl $configFile
            $acl.SetAccessRuleProtection($true, $false)
            foreach ($rule in @($acl.Access)) { [void]$acl.RemoveAccessRule($rule) }
            $me = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            foreach ($identity in @($me, "NT AUTHORITY\SYSTEM")) {
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($identity, "FullControl", "Allow")))
            }
            Set-Acl -Path $configFile -AclObject $acl
            Write-Host "       File access restricted to $me and SYSTEM." -ForegroundColor Yellow
        }
        catch {
            Write-Host "       Could not restrict the file ACL: $_" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " CONFIGURATION CREATED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "File: $configFile" -ForegroundColor Cyan
Write-Host ""
Write-Host "Preview (no email):" -ForegroundColor Yellow
Write-Host "  .\Get-M365Report.ps1 -ConfigFile `"$configFile`" -Preview" -ForegroundColor White
Write-Host ""
Write-Host "Generate PDF:" -ForegroundColor Yellow
Write-Host "  .\Get-M365Report.ps1 -ConfigFile `"$configFile`"" -ForegroundColor White
Write-Host ""
Write-Host "Generate and send by email:" -ForegroundColor Yellow
Write-Host "  .\Get-M365Report.ps1 -ConfigFile `"$configFile`" -SendEmail" -ForegroundColor White
if (-not $msp.Contains('sender_app')) {
    Write-Host ""
    Write-Host "Note: no sender app configured. Email only works if $MSPSender is a mailbox in the" -ForegroundColor Gray
    Write-Host "      client tenant; otherwise add msp.sender_app (see README)." -ForegroundColor Gray
}
Write-Host ""

return $configFile
