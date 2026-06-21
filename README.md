<div align="center">

# M365 Monthly Report

**Monthly Microsoft 365 health and security reports for MSP clients.**

![Status](https://img.shields.io/badge/Status-Active-16A34A?style=flat-square)
![License](https://img.shields.io/badge/License-MIT-2563EB?style=flat-square)
![Casco Digital](https://img.shields.io/badge/Casco-Digital-111827?style=flat-square)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-5391FE?style=flat-square&logo=powershell&logoColor=white)
![Microsoft Graph](https://img.shields.io/badge/Microsoft_Graph-API-00D9FF?style=flat-square&logo=microsoft&logoColor=white)
![PDF](https://img.shields.io/badge/Edge-PDF_Export-0078D4?style=flat-square&logo=microsoftedge&logoColor=white)

</div>

---

Gera relatorios mensais em PDF da saude do Microsoft 365 de clientes MSP. Coleta via Graph API: Secure Score, logins suspeitos, MFA, licencas, storage e uso de servicos.

## Estrutura

```
m365-report/
├── Get-M365Report.ps1          # Script principal
├── Run-Report.ps1              # Wrapper/atalho
├── New-ClientConfig.ps1        # Criar config de novo cliente
├── Schedule-MonthlyReport.ps1  # Agendar no Task Scheduler
├── Setup-Environment.ps1       # Verificar pre-requisitos
├── Fix-SecurityBaseline.ps1    # Correcoes de seguranca (Entra ID)
├── config/
│   └── sample.json             # Exemplo de configuracao
├── templates/
│   ├── report.html             # Template visual do relatorio
│   └── logo.png                # Logo exibido no PDF
├── output/                     # PDFs gerados
└── logs/                       # Logs de execucao
```

## Quick Start

1. Clone o repositorio e rode o setup:
   ```powershell
   git clone https://github.com/cascodigital/m365-report.git
   cd m365-report
   .\Setup-Environment.ps1
   ```

2. Crie a configuracao do cliente:
   ```powershell
   .\New-ClientConfig.ps1 `
     -ClientName "Empresa XYZ" `
     -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
     -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
     -ClientSecret "seu_secret" `
     -Email "ti@empresa.com"
   ```

3. Gere o relatorio:
   ```powershell
   .\Get-M365Report.ps1 -ConfigFile ".\config\EmpresaXYZ.json"
   ```

## Uso

| Comando | Descricao |
|---------|-----------|
| `.\Get-M365Report.ps1 -ConfigFile ".\config\Cliente.json"` | Gera PDF |
| `.\Get-M365Report.ps1 -ConfigFile ".\config\Cliente.json" -Preview` | Abre HTML no navegador |
| `.\Get-M365Report.ps1 -ConfigFile ".\config\Cliente.json" -SendEmail` | Gera PDF e envia por email |
| `.\Run-Report.ps1 -Client NomeCliente` | Atalho — resolve o JSON pelo nome |
| `.\Run-Report.ps1 -All` | Gera relatorio de todos os clientes |
| `.\Run-Report.ps1` | Lista clientes configurados |

## App Registration

No [Entra ID](https://entra.microsoft.com) do tenant do cliente, crie um App Registration com as seguintes permissoes (Application):

| Permissao | Tipo |
|-----------|------|
| `AuditLog.Read.All` | Application |
| `Directory.Read.All` | Application |
| `Reports.Read.All` | Application |
| `SecurityEvents.Read.All` | Application |
| `User.Read.All` | Application |
| `UserAuthenticationMethod.Read.All` | Application |

Para envio de email, o App Registration do **MSP** (nao do cliente) precisa de `Mail.Send` (Application) com admin consent.

## Agendamento

```powershell
.\Schedule-MonthlyReport.ps1 `
  -ConfigFile ".\config\Cliente.json" `
  -Day 30 `
  -Hour "09:00" `
  -SendEmail
```

## Requisitos

- PowerShell 5.1+
- Microsoft Edge (conversao HTML para PDF)
- Modulo `Microsoft.Graph`

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

## Avisos

- **Credenciais:** Nunca commite arquivos `config/*.json` reais — o `.gitignore` ja bloqueia
- **Client Secret:** Secrets expiram — monitore a validade no Entra ID
- **Email:** O envio usa uma Shared Mailbox do MSP, nao do cliente

---

Desenvolvido com 🐢 (e cafe) por **Casco Digital**.
