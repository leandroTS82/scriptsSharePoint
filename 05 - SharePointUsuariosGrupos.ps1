# Conecte-se ao site do SharePoint Online
Connect-PnPOnline -Url "https://butterflygrowth.sharepoint.com/sites/Atrax" -UseWebLogin

# Diretório para salvar os arquivos JSON
$outputDirectory = ".\files\05json"

# Crie o diretório de saída, caso ele não exista
if (!(Test-Path -Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory
}

# Obtenha informações do site
try {
    $site = Get-PnPWeb
    $siteUrl = $site.Url
    $siteName = $site.Title
    $relativeUrl = $siteUrl -replace "https://butterflygrowth.sharepoint.com", ""  # Remove a parte da URL base
    $siteFolderName = if ($relativeUrl -like "*/") { $relativeUrl.TrimEnd('/') } else { $relativeUrl.Split('/')[0] }

    # Salve o nome do site ou subsite em um arquivo .txt
    $siteInfoFilePath = Join-Path -Path $outputDirectory -ChildPath "NomeDoSite.txt"
    $siteName | Out-File -FilePath $siteInfoFilePath -Encoding UTF8
    Write-Output "Arquivo de texto 'NomeDoSite.txt' criado com sucesso com o nome do site: '$siteName'."

    # Obtenha todos os usuários do site
    $users = Get-PnPUser
    $userList = $users | ForEach-Object {
        [PSCustomObject]@{
            Nome  = $_.Title
            Email = $_.Email
            Login = $_.LoginName
        }
    }

    # Converta a lista de usuários para JSON e salve em um arquivo
    $userList | ConvertTo-Json -Depth 3 | Out-File -FilePath "$outputDirectory\05 - SharePointUsuarios.json" -Encoding UTF8
    Write-Output "Arquivo JSON '05 - SharePointUsuarios.json' criado com sucesso."

    # Filtra usuários com email contendo '@' e cria um novo JSON
    $filteredUserList = $userList | Where-Object { $_.Email -match "@" }
    $filteredUserList | ConvertTo-Json -Depth 3 | Out-File -FilePath "$outputDirectory\05 - SharePointUsuarios_ComEmail.json" -Encoding UTF8
    Write-Output "Arquivo JSON '05 - SharePointUsuarios_ComEmail.json' criado com sucesso contendo apenas usuários com email."
}
catch {
    Write-Error "Erro ao obter informações do site ou usuários do SharePoint: $_"
}

# Obtenha todos os grupos do site
try {
    $groups = Get-PnPGroup
    $groupList = $groups | ForEach-Object {
        [PSCustomObject]@{
            Nome        = $_.Title
            Descricao   = $_.Description
        }
    }

    # Converta a lista de grupos para JSON e salve em um arquivo
    $groupList | ConvertTo-Json -Depth 3 | Out-File -FilePath "$outputDirectory\05 - SharePointGrupos.json" -Encoding UTF8
    Write-Output "Arquivo JSON '05 - SharePointGrupos.json' criado com sucesso."
}
catch {
    Write-Error "Erro ao obter grupos do SharePoint: $_"
}
