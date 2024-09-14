# Conecte-se ao site do SharePoint Online
Connect-PnPOnline -Url "https://ltsconsultoria.sharepoint.com" -UseWebLogin

# Obtenha todos os usuários do site
try {
    # Obtenha todos os usuários
    $users = Get-PnPUser

    # Crie uma lista para armazenar os dados dos usuários
    $userList = $users | ForEach-Object {
        [PSCustomObject]@{
            Nome  = $_.Title
            Email = $_.Email
            Login = $_.LoginName
        }
    }

    # Converta a lista de usuários para JSON e salve em um arquivo
    $userList | ConvertTo-Json -Depth 3 | Out-File -FilePath "files/05 - SharePointUsuarios.json" -Encoding UTF8

    Write-Output "Arquivo JSON '05 - SharePointUsuarios.json' criado com sucesso."
}
catch {
    Write-Error "Erro ao obter usuários do SharePoint: $_"
}

# Obtenha todos os grupos do site
try {
    # Obtenha todos os grupos
    $groups = Get-PnPGroup

    # Crie uma lista para armazenar os dados dos grupos
    $groupList = $groups | ForEach-Object {
        [PSCustomObject]@{
            Nome        = $_.Title
            Descricao   = $_.Description
        }
    }

    # Converta a lista de grupos para JSON e salve em um arquivo
    $groupList | ConvertTo-Json -Depth 3 | Out-File -FilePath "files/05 - SharePointGrupos.json" -Encoding UTF8

    Write-Output "Arquivo JSON '05 - SharePointGrupos.json' criado com sucesso."
}
catch {
    Write-Error "Erro ao obter grupos do SharePoint: $_"
}
