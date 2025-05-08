# Conecte-se ao site do SharePoint Online
Connect-PnPOnline -Url "https://butterflygrowth.sharepoint.com" -UseWebLogin

# Diretório para salvar os arquivos JSON
$outputDirectory = ".\files\07json"

# Crie o diretório de saída, caso ele não exista
if (!(Test-Path -Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory
}

# Obtenha informações da coleção de sites
try {
    # Obtenha todos os sites da coleção
    $sites = Get-PnPTenantSite

    # Crie um array para armazenar os dados dos sites
    $siteList = @()

    # Preencha o array com os dados dos sites
    foreach ($site in $sites) {
        $siteList += [PSCustomObject]@{
            Nome = $site.Title
            Url  = $site.Url
        }
    }

    # Converta a lista de sites para JSON e salve em um arquivo
    $jsonFilePath = Join-Path -Path $outputDirectory -ChildPath "07 - SharePointSites.json"
    $siteList | ConvertTo-Json -Depth 3 | Out-File -FilePath $jsonFilePath -Encoding UTF8
    Write-Output "Arquivo JSON '07 - SharePointSites.json' criado com sucesso contendo todos os sites da coleção."

} catch {
    Write-Error "Erro ao obter informações da coleção de sites do SharePoint: $_"
}
