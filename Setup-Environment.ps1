<#
.SYNOPSIS
    Checks the environment for running the M365 reports and prints the
    app registration setup steps.

.DESCRIPTION
    - Checks prerequisites (PowerShell version, Microsoft Edge for PDF export)
    - Creates the working folders (config, output, logs, templates)
    - Prints the required Microsoft Graph application permissions

    No PowerShell modules are installed or required: the report uses raw
    Microsoft Graph REST calls.
#>

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " M365 MONTHLY REPORT - SETUP" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# PowerShell version
$psVersion = $PSVersionTable.PSVersion
if ($psVersion.Major -lt 5 -or ($psVersion.Major -eq 5 -and $psVersion.Minor -lt 1)) {
    Write-Host "[WARN] PowerShell $($psVersion.Major).$($psVersion.Minor) detected. PowerShell 5.1 or later is required." -ForegroundColor Yellow
}
else {
    Write-Host "[CHECK] PowerShell version: $($psVersion.Major).$($psVersion.Minor)" -ForegroundColor Green
}

# Microsoft Edge (HTML -> PDF)
$edgePaths = @(
    "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
)
$edgeFound = $false
foreach ($path in $edgePaths) {
    if (Test-Path $path) {
        Write-Host "[CHECK] Microsoft Edge found: $path" -ForegroundColor Green
        $edgeFound = $true
        break
    }
}
if (!$edgeFound) {
    Write-Host "[WARN] Microsoft Edge not found. It is required to generate PDFs (use -Preview for HTML only)." -ForegroundColor Yellow
}

# Folder structure
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$folders = @("config", "output", "logs", "templates")
foreach ($folder in $folders) {
    $folderPath = Join-Path $ScriptPath $folder
    if (!(Test-Path $folderPath)) {
        New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
        Write-Host "[CREATED] Folder created: $folder" -ForegroundColor Green
    }
    else {
        Write-Host "[OK] Folder exists: $folder" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host " APP REGISTRATION (CLIENT TENANT)" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "In each client tenant, create an app registration:"
Write-Host ""
Write-Host "1. Go to: https://entra.microsoft.com" -ForegroundColor Cyan
Write-Host "2. Identity > Applications > App registrations"
Write-Host "3. + New registration"
Write-Host "4. Name: 'MSP Monthly Report' (or similar)"
Write-Host "5. Supported account types: 'Single tenant'"
Write-Host "6. Redirect URI: leave empty"
Write-Host "7. Register"
Write-Host ""
Write-Host "Then add these API PERMISSIONS:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Microsoft Graph (Application permissions):" -ForegroundColor Cyan
Write-Host "  - User.Read.All                     (users)"
Write-Host "  - Directory.Read.All                (admin roles, groups, domains, licenses)"
Write-Host "  - AuditLog.Read.All                 (sign-in logs, MFA registration report)"
Write-Host "  - UserAuthenticationMethod.Read.All (MFA fallback when the registration report is unavailable)"
Write-Host "  - Reports.Read.All                  (usage reports: storage, mailboxes, service adoption)"
Write-Host "  - SecurityEvents.Read.All           (Secure Score)"
Write-Host "  - SecurityAlert.Read.All            (security alerts)"
Write-Host "  - ServiceHealth.Read.All            (Microsoft 365 service health)"
Write-Host ""
Write-Host "  Sign-in logs require an Entra ID P1/P2 license in the client tenant." -ForegroundColor Gray
Write-Host "  Any missing permission/license shows as 'Not collected' in the report." -ForegroundColor Gray
Write-Host ""
Write-Host "IMPORTANT: click 'Grant admin consent' after adding the permissions." -ForegroundColor Red
Write-Host ""
Write-Host "Credentials - choose one:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  A) Certificate (recommended, nothing secret stored in the config):" -ForegroundColor Cyan
Write-Host '     $cert = New-SelfSignedCertificate -Subject "CN=MSP Monthly Report" -CertStoreLocation Cert:\CurrentUser\My -KeySpec Signature -KeyExportPolicy NonExportable -NotAfter (Get-Date).AddYears(2)'
Write-Host '     Export-Certificate -Cert $cert -FilePath .\msp-report.cer'
Write-Host "     Upload msp-report.cer under 'Certificates & secrets > Certificates' and use"
Write-Host '     $cert.Thumbprint as -CertificateThumbprint.'
Write-Host ""
Write-Host "  B) Client secret:" -ForegroundColor Cyan
Write-Host "     'Certificates & secrets' > + New client secret > copy the VALUE (shown once)."
Write-Host "     The secret is stored in the client config file - keep that file protected."
Write-Host ""
Write-Host "========================================" -ForegroundColor Yellow
Write-Host " EMAIL SENDER (MSP TENANT)" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "Reports are sent from a mailbox in YOUR (MSP) tenant, e.g. reports@yourmsp.com:"
Write-Host ""
Write-Host "1. Create (or reuse) a shared mailbox in the Exchange admin center."
Write-Host "2. Create an app registration in YOUR tenant with Mail.Send (Application) + admin consent."
Write-Host "3. Recommended: restrict that app to the sender mailbox (Exchange RBAC for Applications"
Write-Host "   or an application access policy)."
Write-Host "4. Put that app in the client config under msp.sender_app"
Write-Host "   (tenant_id, client_id, and certificate_thumbprint or client_secret)."
Write-Host ""
Write-Host "If the sender mailbox is in the client tenant itself, msp.sender_app is not needed;"
Write-Host "the client app then needs Mail.Send instead."
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " NEXT STEP" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "After creating the app registration, run:" -ForegroundColor Cyan
Write-Host ""
Write-Host '.\New-ClientConfig.ps1 -ClientName "Contoso" -TenantId "<tenant-id>" -ClientId "<client-id>" -CertificateThumbprint "<thumbprint>" -Email "it@contoso.com"'
Write-Host ""
Write-Host "Or copy and edit: config\sample.json"
Write-Host ""
