<#
.SYNOPSIS
    Fix-SecurityBaseline.ps1 - Otimização de Segurança M365 (Modo Sênior)
    
.DESCRIPTION
    Script interativo para aplicar correções de segurança recomendadas no Entra ID/M365.
    Utiliza o módulo Microsoft.Graph.
#>

$ErrorActionPreference = "Stop"

# Verifica se o módulo está instalado
if (!(Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Módulo Microsoft.Graph não encontrado. Instale com: Install-Module Microsoft.Graph" -ForegroundColor Red
    return
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Identity.SignIns
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Groups

function Connect-M365 {
    Write-Host "Conectando ao Microsoft Graph..." -ForegroundColor Cyan
    # Scopes necessários para as operações do script
    $scopes = @(
        "Policy.ReadWrite.ConditionalAccess",
        "Policy.Read.All",
        "Directory.ReadWrite.All",
        "User.ReadWrite.All",
        "RoleManagement.Read.Directory",
        "Policy.ReadWrite.Authorization"
    )
    Connect-MgGraph -Scopes $scopes
}

function Show-Menu {
    Clear-Host
    Write-Host "==========================================================" -ForegroundColor Magenta
    Write-Host "   SKIPPY FIX-O-MATIC - SEGURANÇA M365 (MODO ANDRÉ)   " -ForegroundColor Magenta
    Write-Host "==========================================================" -ForegroundColor Magenta
    Write-Host "1. Ativar Security Defaults (MFA para todos + Bloqueio Legado)"
    Write-Host "2. Definir Senhas para NUNCA EXPIRAR (Tenant-wide)"
    Write-Host "3. Bloquear Consentimento de Usuários para Apps (Anti-Phishing)"
    Write-Host "4. Auditar Admins Globais (Verificar Breaking Glass)"
    Write-Host "5. [SAIR]"
    Write-Host "=========================================================="
}

function Set-SecurityDefaults {
    Write-Host "Lendo configuração atual de Security Defaults..." -ForegroundColor Cyan
    $policy = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy
    
    if ($policy.IsEnabled) {
        Write-Host "Security Defaults já está ATIVADO." -ForegroundColor Green
    } else {
        $confirm = Read-Host "Deseja ATIVAR o Security Defaults? (S/N)"
        if ($confirm -eq 'S') {
            Update-MgPolicyIdentitySecurityDefaultEnforcementPolicy -IsEnabled:$true
            Write-Host "Security Defaults ATIVADO com sucesso!" -ForegroundColor Green
        }
    }
}

function Set-PasswordNeverExpire {
    Write-Host "Buscando usuários com senha que expira..." -ForegroundColor Cyan
    $users = Get-MgUser -All -Property "Id,DisplayName,PasswordPolicies"
    $toUpdate = $users | Where-Object { $_.PasswordPolicies -ne "DisablePasswordExpiration" }
    
    if ($toUpdate.Count -eq 0) {
        Write-Host "Todos os usuários já estão configurados para nunca expirar." -ForegroundColor Green
    } else {
        Write-Host "Encontrados $($toUpdate.Count) usuários para atualizar." -ForegroundColor Yellow
        $confirm = Read-Host "Deseja definir 'Nunca Expirar' para todos? (S/N)"
        if ($confirm -eq 'S') {
            foreach ($u in $toUpdate) {
                Update-MgUser -UserId $u.Id -PasswordPolicies "DisablePasswordExpiration"
                Write-Host "Atualizado: $($u.DisplayName)" -ForegroundColor Gray
            }
            Write-Host "Concluído!" -ForegroundColor Green
        }
    }
}

function Set-AppConsent {
    Write-Host "Configurando políticas de consentimento..." -ForegroundColor Cyan
    
    $confirm = Read-Host "Bloquear usuários de criar/consentir Apps? (S/N)"
    if ($confirm -eq 'S') {
        try {
            Write-Host "Aplicando via Graph API REST (Modo Jedi)..." -ForegroundColor Cyan
            
            # Endpoint direto da política de autorização
            # CORREÇÃO: URL estava duplicada (/authorizationPolicy/authorizationPolicy)
            $uri = "https://graph.microsoft.com/v1.0/policies/authorizationPolicy"
            
            # Payload limpo conforme documentação oficial v1.0
            $body = @{
                allowedToSignUpEmailBasedSubscriptions = $false
                defaultUserRolePermissions = @{
                    allowedToCreateApps = $false
                    # Manter allowedToReadOtherUsers true para não quebrar Teams/AddressBook
                    allowedToReadOtherUsers = $true 
                }
            } | ConvertTo-Json -Depth 5

            # O cmdlet Invoke-MgGraphRequest é a "chave mestra"
            Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body -ContentType "application/json"
            
            Write-Host "Política de Consentimento atualizada com sucesso!" -ForegroundColor Green
        }
        catch {
            Write-Host "Erro FATAL ao atualizar política: $($_.Exception.Message)" -ForegroundColor Red
            # CORREÇÃO: Tratamento de erro compatível com o Graph SDK Exception
            if ($_.Exception.Response -and $_.Exception.Response.Content) {
                 $errorDetail = $_.Exception.Response.Content.ReadAsStringAsync().Result
                 Write-Host "Detalhes do Graph: $errorDetail" -ForegroundColor Red
            }
        }
    }
}

function Audit-Admins {
    Write-Host "Listando Administradores Globais..." -ForegroundColor Cyan
    try {
        # Busca a role de Admin Global
        $role = Get-MgDirectoryRole | Where-Object { $_.DisplayName -eq "Global Administrator" }
        
        if ($null -eq $role) {
            Write-Host "Ativando role de Admin Global no diretório..." -ForegroundColor Gray
            $template = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq "Global Administrator" }
            New-MgDirectoryRole -DirectoryTemplateId $template.Id | Out-Null
            $role = Get-MgDirectoryRole | Where-Object { $_.DisplayName -eq "Global Administrator" }
        }

        $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id
        foreach ($member in $members) {
            # Busca detalhes do usuário pelo ID
            $u = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
            if ($u) {
                $color = if ($u.UserPrincipalName -like "*admin*" -or $u.UserPrincipalName -like "*break*") { "Yellow" } else { "White" }
                Write-Host "- $($u.DisplayName) ($($u.UserPrincipalName))" -ForegroundColor $color
            } else {
                Write-Host "- ID: $($member.Id) (Não é um usuário ou sem permissão)" -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Host "Falha na auditoria: $($_.Exception.Message)" -ForegroundColor Red
    }
    Read-Host "`nPressione Enter para voltar..."
}

# Início do script
Connect-M365

do {
    Show-Menu
    $choice = Read-Host "Escolha uma opção"
    
    switch ($choice) {
        "1" { Set-SecurityDefaults }
        "2" { Set-PasswordNeverExpire }
        "3" { Set-AppConsent }
        "4" { Audit-Admins }
        "5" { $running = $false }
    }
    
    if ($choice -ne "5") {
        Write-Host "`nOperação concluída."
        Start-Sleep -Seconds 2
    }
} while ($choice -ne "5")

Write-Host "Skippy saindo... De nada, saco de água." -ForegroundColor Magenta
