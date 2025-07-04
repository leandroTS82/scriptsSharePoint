# Usado para adicionar funções a um usuário no Azure sharepoint

<#
.SYNOPSIS
    Adiciona ou remove uma função do Azure AD para um usuário externo.

.PARAMETER UserUPN
    UPN do usuário (ex: leandrofire_live.com#EXT#@dominio.onmicrosoft.com)

.PARAMETER RoleName
    Nome da função do Azure (ex: Global Administrator)

.PARAMETER Action
    A para adicionar ou R para remover

.EXAMPLE
    .\Manage-AzureADRole.ps1 -UserUPN "leandrofire_live.com#EXT#@ButterflyGrowth.onmicrosoft.com" -RoleName "Global Administrator" -Action A
#>

param (
    [string]$UserUPN,
    [string]$RoleName,
    [ValidateSet("A", "R")]
    [string]$Action
)

# Instala e importa módulo
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph -Force
Import-Module Microsoft.Graph.RoleManagement.Directory -Force

# Conecta ao Microsoft Graph
Connect-MgGraph -Scopes "RoleManagement.ReadWrite.Directory", "Directory.Read.All"

# Seleciona o usuário
if (-not $UserUPN) {
    $externalUsers = Get-MgUser -All | Where-Object { $_.UserPrincipalName -like "*#EXT#*" } |
        Select-Object DisplayName, UserPrincipalName, Id
    $selectedUser = $externalUsers | Out-GridView -Title "Selecione um usuário externo" -PassThru
    if (-not $selectedUser) { Write-Host "❌ Nenhum usuário selecionado."; return }
    $user = $selectedUser
} else {
    $user = Get-MgUser -All | Where-Object { $_.UserPrincipalName -eq $UserUPN }
    if (-not $user) { Write-Host "❌ Usuário '$UserUPN' não encontrado."; return }
}

# Seleciona a função
if (-not $RoleName) {
    $availableRoles = Get-MgRoleManagementDirectoryRoleDefinition | Sort-Object DisplayName
    $selectedRole = $availableRoles | Out-GridView -Title "Selecione a função desejada" -PassThru
    if (-not $selectedRole) { Write-Host "❌ Nenhuma função selecionada."; return }
    $role = $selectedRole
} else {
    $role = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -eq $RoleName }
    if (-not $role) { Write-Host "❌ Função '$RoleName' não encontrada."; return }
}

# Define ação se não informada
if (-not $Action) {
    $Action = Read-Host "Digite 'A' para Adicionar ou 'R' para Remover a função"
}

# Execução
switch ($Action.ToUpper()) {
    "A" {
        try {
            $assignment = New-MgRoleManagementDirectoryRoleAssignment `
                -DirectoryScopeId "/" `
                -PrincipalId $user.Id `
                -RoleDefinitionId $role.Id

            Write-Host "✅ Função '$($role.DisplayName)' atribuída ao usuário '$($user.DisplayName)'." -ForegroundColor Green
            Write-Host "AssignmentId: $($assignment.Id)"
            
            # Salvar ID em arquivo
            $logPath = ".\AzureRoleAssignments.log"
            Add-Content -Path $logPath -Value "$(Get-Date -Format s) | Adicionado | User: $($user.UserPrincipalName) | Role: $($role.DisplayName) | AssignmentId: $($assignment.Id)"
        } catch {
            Write-Error "Erro ao atribuir função: $_"
        }
    }

    "R" {
        try {
            $existingAssignments = Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object {
                $_.PrincipalId -eq $user.Id -and $_.RoleDefinitionId -eq $role.Id -and $_.DirectoryScopeId -eq "/"
            }

            if ($existingAssignments) {
                foreach ($assignment in $existingAssignments) {
                    Remove-MgRoleManagementDirectoryRoleAssignment -Id $assignment.Id -Confirm:$false
                    Write-Host "✅ Função '$($role.DisplayName)' removida do usuário '$($user.DisplayName)'." -ForegroundColor Yellow
                    $logPath = ".\AzureRoleAssignments.log"
                    Add-Content -Path $logPath -Value "$(Get-Date -Format s) | Removido  | User: $($user.UserPrincipalName) | Role: $($role.DisplayName) | AssignmentId: $($assignment.Id)"
                }
            } else {
                Write-Host "⚠️ Nenhuma função atribuída encontrada para esse usuário."
            }
        } catch {
            Write-Error "Erro ao remover função: $_"
        }
    }

    Default {
        Write-Host "❌ Ação inválida. Use 'A' para adicionar ou 'R' para remover." -ForegroundColor Red
    }
}
