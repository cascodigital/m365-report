<div align="center">

# M365 Monthly Report

**Monthly Microsoft 365 health and security reports for MSP clients.**

![Status](https://img.shields.io/badge/Status-Active-16A34A?style=flat-square)
![License](https://img.shields.io/badge/License-MIT-2563EB?style=flat-square)
![Casco Digital](https://img.shields.io/badge/Casco-Digital-111827?style=flat-square)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-5391FE?style=flat-square&logo=powershell&logoColor=white)
![Microsoft Graph](https://img.shields.io/badge/Microsoft_Graph-REST-00D9FF?style=flat-square&logo=microsoft&logoColor=white)
![PDF](https://img.shields.io/badge/Edge-PDF_Export-0078D4?style=flat-square&logo=microsoftedge&logoColor=white)

</div>

---

Builds a monthly Microsoft 365 health report (HTML + PDF) for each client tenant and optionally emails it. Data is collected with plain Microsoft Graph REST calls from an app registration (client credentials). No PowerShell modules are needed.

The report never shows a number that was not measured. If a permission or license is missing, that item shows **"Not collected"** and the summary and recommendations say so. The report does not fill the gap with a placeholder value.

## Structure

```
m365-report/
├── Get-M365Report.ps1          # Main script (collect, render, PDF, email)
├── Run-Report.ps1              # Shortcut: run by client name or for all clients
├── New-ClientConfig.ps1        # Create a client config file
├── Schedule-MonthlyReport.ps1  # Register a monthly Task Scheduler task
├── Setup-Environment.ps1       # Check prerequisites, print setup steps
├── config/
│   └── sample.json             # Example configuration
├── lang/
│   └── pt-BR.json              # Brazilian Portuguese strings (English is built in)
├── templates/
│   ├── report.html             # Report template
│   └── logo.png                # Default header logo (override with msp.logo)
├── output/                     # Generated reports
└── logs/                       # Run logs
```

## What is collected

| Section | Source (Graph) | Notes |
|---|---|---|
| Secure Score + change vs ~30 days ago | `/security/secureScores` | |
| Enabled member users | `/users` (paged) | Guests excluded |
| MFA registration per user | `/reports/authenticationMethods/userRegistrationDetails` (paged) | Falls back to `/users/{id}/authentication/methods` if the report is unavailable. "Registered" is not the same as "enforced". |
| Admins with MFA | `/directoryRoles` + members, joined with the MFA data | Active members of roles whose name contains "Administrator". PIM-eligible (not activated) assignments and group-based role members are not counted. |
| Sign-ins, last 30 days | `/auditLogs/signIns` (paged, capped by `options.max_signin_pages`, default 50 × 1000) | Requires Entra ID P1/P2. Interrupted sign-ins (e.g. MFA prompts) count as failed. Marked with `+` when the cap is reached. |
| Security alerts, last 30 days | `/security/alerts_v2` | Total, high severity, unresolved |
| Service health | `/admin/serviceAnnouncement/healthOverviews` and `/issues` | Current status + incidents/advisories started in the last 30 days |
| Licenses | `/subscribedSkus` | Free/trial SKUs ignored |
| Microsoft 365 groups / Teams | `/groups` (paged) | |
| Storage, mailboxes, service adoption | Usage reports (`/reports/get...`) | See below |

**Usage reports.** v1.0 is used first: it returns CSV through a redirect, and the script parses it. `$format=application/json` exists only on `/beta`, so beta is used only as a fallback when the v1.0 CSV cannot be read. Storage is the latest daily snapshot, not a sum over the period. Shared mailboxes are counted only when the mailbox report includes a recipient type. Otherwise the report shows "Not collected".

**Not collected (by design):** malware/phishing counters, uptime percentages, and extrapolated numbers. Earlier versions showed hardcoded values for these. They were removed.

## Required permissions

Client tenant app registration, **Application** permissions with admin consent:

| Permission | Used for |
|---|---|
| `User.Read.All` | Users |
| `Directory.Read.All` | Admin roles, groups, domains, licenses |
| `AuditLog.Read.All` | Sign-in logs, MFA registration report |
| `UserAuthenticationMethod.Read.All` | MFA fallback (per-user methods) |
| `Reports.Read.All` | Usage reports |
| `SecurityEvents.Read.All` | Secure Score |
| `SecurityAlert.Read.All` | Security alerts |
| `ServiceHealth.Read.All` | Service health |

Sender app (MSP tenant, only for `-SendEmail`): `Mail.Send` (Application). Restricting it to the sender mailbox with Exchange RBAC for Applications or an application access policy is strongly recommended. If the sender mailbox is in the client tenant itself, no sender app is needed, but the client app then needs `Mail.Send`.

## Authentication

Each app (client and sender) can use either:

- **Certificate** (recommended): `certificate_thumbprint`. The certificate with its private key must be in `CurrentUser\My` or `LocalMachine\My`. No secret is stored in the config.
- **Client secret**: `client_secret`. It is stored in plain text in the config file. `New-ClientConfig.ps1` warns about this and restricts the file ACL to the current user and SYSTEM.

## Quick start

1. Check prerequisites and read the setup steps:
   ```powershell
   .\Setup-Environment.ps1
   ```

2. Create the client config:
   ```powershell
   .\New-ClientConfig.ps1 `
     -ClientName "Contoso" `
     -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
     -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
     -CertificateThumbprint "<thumbprint>" `
     -Email "it@contoso.com"
   ```
   The script prompts for MSP name, contact and sender if they are not passed. For email from your own tenant, add `-SenderTenantId`, `-SenderClientId` and `-SenderCertificateThumbprint` (or `-SenderClientSecret`).

3. Generate the report:
   ```powershell
   .\Get-M365Report.ps1 -ConfigFile ".\config\contoso.json" -Preview
   ```

## Usage

| Command | Description |
|---------|-------------|
| `.\Get-M365Report.ps1 -ConfigFile ".\config\contoso.json"` | Generate PDF |
| `.\Get-M365Report.ps1 -ConfigFile ".\config\contoso.json" -Preview` | Open the HTML in the browser |
| `.\Get-M365Report.ps1 -ConfigFile ".\config\contoso.json" -SendEmail` | Generate PDF and send by email |
| `.\Run-Report.ps1 -Client contoso` | Shortcut, resolves `config\contoso.json` |
| `.\Run-Report.ps1 -All` | All clients (skips `sample.json`) |
| `.\Run-Report.ps1` | List configured clients |

## Configuration

See [`config/sample.json`](config/sample.json). Main keys:

| Key | Description |
|---|---|
| `client.name`, `client.tenant_id`, `client.client_id` | Client tenant app |
| `client.certificate_thumbprint` / `client.client_secret` | Client app credential |
| `msp.name`, `msp.contact`, `msp.tagline` | Branding (header/footer). The tagline is optional. |
| `msp.logo` | Logo path (absolute, or relative to the config file or the script folder). Defaults to `templates/logo.png` if that file exists. |
| `msp.sender` | Mailbox the report is sent from |
| `msp.sender_app.{tenant_id, client_id, certificate_thumbprint \| client_secret}` | MSP tenant app used to send email |
| `email.recipients`, `email.cc`, `email.subject_template` | Email. `{{PERIOD}}` is replaced with the report month. |
| `options.language` | `en` (default) or `pt-BR` |
| `options.include_users_without_mfa`, `options.include_recommendations` | Report sections (default `true`) |
| `options.max_signin_pages` | Sign-in paging cap (default 50 pages of 1000) |

**Backward compatibility:** configs that use the original Portuguese keys still work: `cliente.nome/tenant_id/client_id/client_secret`, `msp.nome/contato/remetente`, `email.destinatarios/assunto_template` (`{{PERIODO}}`), and `opcoes.incluir_usuarios_sem_mfa/incluir_recomendacoes/idioma`. English keys win when both are present. A legacy config with `"idioma": "pt-BR"` keeps producing a Portuguese report.

## Scheduling

```powershell
.\Schedule-MonthlyReport.ps1 `
  -ConfigFile ".\config\contoso.json" `
  -Day 1 `
  -Hour "09:00" `
  -SendEmail
```

The task runs on day 1 by default, since the report covers the previous month. It uses `pwsh.exe` when PowerShell 7 is installed and falls back to `powershell.exe`. The config path is stored as an absolute path.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Microsoft Edge (HTML to PDF). Use `-Preview` for HTML only.
- No PowerShell modules

## Notes

- **Credentials:** never commit real `config/*.json` files. `.gitignore` already excludes them, except `sample.json`.
- **Secrets/certificates expire.** Track their expiry in Entra ID.
- **Email** goes out from the MSP mailbox through the MSP tenant sender app. The client app is used only when the sender mailbox is in the client tenant. Otherwise the script stops with a clear error instead of a 403.

## License

[MIT](LICENSE)

---

Built by **Casco Digital**.
