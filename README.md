# m365-report

A PowerShell tool that builds monthly Microsoft 365 health and security reports (HTML + PDF) for MSP client tenants and optionally emails them.

![License](https://img.shields.io/badge/license-MIT-blue) ![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE) ![Status](https://img.shields.io/badge/status-active-green)

## Overview

MSPs need to show each client, every month, where their Microsoft 365 tenant stands: Secure Score, MFA
coverage, sign-ins, alerts, licenses and usage. The admin centers spread this across many blades, and most
reporting products require extra licensing or agents. This script collects the data with plain Microsoft Graph
REST calls from an app registration (client credentials, no PowerShell modules), renders a branded report and
exports it to PDF with headless Edge.

The report never shows a number that was not measured. If a permission or license is missing, that item shows
**"Not collected"** and the summary and recommendations say so instead of filling the gap with a placeholder.

## Features

- Collects Secure Score, users, MFA registration, admin MFA, sign-ins, alerts, service health, licenses,
  groups/Teams, storage and service adoption from Microsoft Graph.
- Authenticates per client tenant with a certificate (recommended) or a client secret.
- Renders an HTML report from `templates/report.html` with MSP name, contact, tagline and logo.
- Exports to PDF through headless Microsoft Edge, or opens the HTML with `-Preview`.
- Emails the PDF through Graph `sendMail` from a sender app in the MSP tenant.
- Generates rule-based recommendations (admins without MFA, low Secure Score, open alerts, failed sign-ins).
- Produces English or Brazilian Portuguese reports (`lang/pt-BR.json`).
- Runs for one client, all clients, or on a monthly Task Scheduler schedule.

## How it works

```text
config/<client>.json
        |
        v
Get-M365Report.ps1 --(client credentials)--> Microsoft Graph (client tenant)
        |  collect -> compute -> fill templates/report.html (+ language strings)
        v
output/*.html --(msedge --headless --print-to-pdf)--> output/*.pdf
        |
        v  (-SendEmail)
sender app token (MSP tenant) --> /users/<sender>/sendMail
```

`Run-Report.ps1` and `Schedule-MonthlyReport.ps1` are thin wrappers around `Get-M365Report.ps1`.
`New-ClientConfig.ps1` writes the per-client JSON, and `Setup-Environment.ps1` checks prerequisites and prints
the app-registration steps.

What is collected:

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
| Storage, mailboxes, service adoption | Usage reports (`/reports/get...`) | See the note below |

Usage reports: v1.0 is used first. It returns CSV through a redirect, and the script parses it. `$format=application/json` exists only on `/beta`, so beta is used only as a fallback when the v1.0 CSV cannot be read. Storage is the latest daily snapshot, not a sum over the period. Shared mailboxes are counted only when the mailbox report includes a recipient type. Otherwise the report shows "Not collected".

Not collected by design: malware/phishing counters, uptime percentages and extrapolated numbers.

## Requirements

- Windows with Windows PowerShell 5.1 or PowerShell 7+.
- Microsoft Edge (for HTML to PDF). Use `-Preview` for HTML only.
- No PowerShell modules.
- An app registration in each client tenant with these **Application** permissions and admin consent:

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

Each app (client and sender) can authenticate with:

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

## Configuration

Per-client JSON in `config/` (see [`config/sample.json`](config/sample.json)). Real configs are git-ignored;
only `sample.json` is tracked.

| Variable / key | Required | Default | Description |
|---|---|---|---|
| `client.name`, `client.tenant_id`, `client.client_id` | Yes | — | Client tenant app |
| `client.certificate_thumbprint` / `client.client_secret` | Yes (one) | — | Client app credential |
| `msp.name`, `msp.contact`, `msp.tagline` | No | — | Branding (header/footer). The tagline is optional. |
| `msp.logo` | No | `templates/logo.png` | Logo path (absolute, or relative to the config file or the script folder). Defaults to `templates/logo.png` if that file exists. |
| `msp.sender` | For `-SendEmail` | — | Mailbox the report is sent from |
| `msp.sender_app.{tenant_id, client_id, certificate_thumbprint \| client_secret}` | No | — | MSP tenant app used to send email |
| `email.recipients`, `email.cc`, `email.subject_template` | For `-SendEmail` | — | Email. `{{PERIOD}}` is replaced with the report month. |
| `options.language` | No | `en` | `en` (default) or `pt-BR` |
| `options.include_users_without_mfa`, `options.include_recommendations` | No | `true` | Report sections (default `true`) |
| `options.max_signin_pages` | No | `50` | Sign-in paging cap (default 50 pages of 1000) |

**Backward compatibility:** configs that use the original Portuguese keys still work: `cliente.nome/tenant_id/client_id/client_secret`, `msp.nome/contato/remetente`, `email.destinatarios/assunto_template` (`{{PERIODO}}`), and `opcoes.incluir_usuarios_sem_mfa/incluir_recomendacoes/idioma`. English keys win when both are present. A legacy config with `"idioma": "pt-BR"` keeps producing a Portuguese report.

## Usage

| Command | Description |
|---|---|
| `.\Get-M365Report.ps1 -ConfigFile ".\config\contoso.json"` | Generate PDF |
| `.\Get-M365Report.ps1 -ConfigFile ".\config\contoso.json" -Preview` | Open the HTML in the browser |
| `.\Get-M365Report.ps1 -ConfigFile ".\config\contoso.json" -SendEmail` | Generate PDF and send by email |
| `.\Run-Report.ps1 -Client contoso` | Shortcut, resolves `config\contoso.json` |
| `.\Run-Report.ps1 -All` | All clients (skips `sample.json`) |
| `.\Run-Report.ps1` | List configured clients |
| `.\Run-Report.ps1 -Client contoso -Preview` | Shortcut with HTML preview (`-SendEmail` is also passed through) |

Monthly schedule (Windows Task Scheduler):

```powershell
.\Schedule-MonthlyReport.ps1 `
  -ConfigFile ".\config\contoso.json" `
  -Day 1 `
  -Hour "09:00" `
  -SendEmail
```

The task runs on day 1 by default, since the report covers the previous month. It uses `pwsh.exe` when PowerShell 7 is installed and falls back to `powershell.exe`. The config path is stored as an absolute path.

## Limitations

- "MFA registered" is not the same as "MFA enforced"; the report measures registration.
- Admin MFA counts active members of roles whose name contains "Administrator"; PIM-eligible and group-based
  role assignments are not counted.
- Sign-in data requires Entra ID P1/P2 and is capped by `options.max_signin_pages`.
- Client secrets, when used, are stored in plain text in the config file. Secrets and certificates expire;
  track them in Entra ID.
- The scheduled task uses S4U logon, which cannot read `CurrentUser` certificate private keys; use a
  certificate in `LocalMachine\My` for scheduled runs.
- Windows only (Edge path, Task Scheduler, certificate store).

## License

MIT — see [LICENSE](LICENSE).
