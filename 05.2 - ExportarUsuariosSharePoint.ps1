# Defina o diretório para salvar os arquivos JSON e CSV
$outputDirectoryJson = ".\files\07json"
$outputDirectoryCsv = ".\files\07csv"

# Crie o diretório de saída para os arquivos CSV, caso não exista
if (!(Test-Path -Path $outputDirectoryCsv)) {
    New-Item -ItemType Directory -Path $outputDirectoryCsv
}

# Caminho para o arquivo JSON
$jsonFilePath = Join-Path -Path $outputDirectoryJson -ChildPath "07 - SharePointSites.json"

# Verifique se o arquivo JSON existe
if (-Not (Test-Path -Path $jsonFilePath)) {
    Write-Error "O arquivo JSON não foi encontrado: $jsonFilePath"
    exit
}

# Leia o conteúdo do arquivo JSON
$jsonContent = Get-Content -Path $jsonFilePath -Raw | ConvertFrom-Json

# Itere sobre cada site no JSON
foreach ($site in $jsonContent) {
    $siteName = $site.Nome
    $siteUrl = $site.Url

    # Conecte-se ao site do SharePoint Online
    Connect-PnPOnline -Url $siteUrl -UseWebLogin

    # Obtenha todos os usuários do site
    try {
        $users = Get-PnPUser
        $userList = $users | Where-Object { $_.Email -match "@" } | ForEach-Object {
            [PSCustomObject]@{
                Nome  = $_.Title
                Email = $_.Email
                Login = $_.LoginName
            }
        }

        # Verifique se há usuários a serem exportados
        if ($userList.Count -eq 0) {
            Write-Output "Nenhum usuário com email encontrado para o site '$siteName'."
            continue
        }

        # Caminho completo para o arquivo CSV
        $csvFilePath = Join-Path -Path $outputDirectoryCsv -ChildPath ($siteName + ".csv")

        # Exporta os dados para um arquivo CSV
        $userList | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
        Write-Output "Arquivo CSV '$csvFilePath' criado com sucesso para o site '$siteName'."
    } catch {
        Write-Error "Erro ao obter usuários do site '$siteName': $_"
    }
}
