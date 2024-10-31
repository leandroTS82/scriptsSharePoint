# Conecte-se ao site do SharePoint Online
Connect-PnPOnline -Url "https://butterflygrowth.sharepoint.com" -UseWebLogin

# Diretório para salvar os arquivos JSON
$outputDirectory = ".\files\07json"

# Crie o diretório de saída, caso ele não exista
if (!(Test-Path -Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory
}

# Defina os termos para exclusão
$excludedTerms = @("", "RedirectSite", "search")

# Obtenha informações da coleção de sites
try {
    # Obtenha todos os sites da coleção
    $sites = Get-PnPTenantSite

    # Crie arrays para armazenar os dados dos sites incluídos e excluídos
    $siteList = @()
    $excludedSites = @()

    # Preencha o array com os dados dos sites
    foreach ($site in $sites) {
        $siteObj = [PSCustomObject]@{
            Nome = $site.Title
            Url  = $site.Url
        }
        
        # Verifique se o Nome ou Url devem ser excluídos
        $shouldExclude = ($site.Title -eq "" -or $site.Title -eq "RedirectSite" -or $site.Url -like "*search*")
        
        # Adiciona o site no array correto
        if ($shouldExclude) {
            $excludedSites += $siteObj
        } else {
            $siteList += $siteObj
        }
    }

    # Converta a lista de sites incluídos para JSON e salve em um arquivo
    $jsonFilePath = Join-Path -Path $outputDirectory -ChildPath "07 - SharePointSites.json"
    $siteList | ConvertTo-Json -Depth 3 | Out-File -FilePath $jsonFilePath -Encoding UTF8
    Write-Output "Arquivo JSON '07 - SharePointSites.json' criado com sucesso contendo todos os sites incluídos da coleção."

    # Converta a lista de sites excluídos para JSON e salve em um arquivo
    $excludedJsonFilePath = Join-Path -Path $outputDirectory -ChildPath "07 - SharePointSitesExcluidos.json"
    $excludedSites | ConvertTo-Json -Depth 3 | Out-File -FilePath $excludedJsonFilePath -Encoding UTF8
    Write-Output "Arquivo JSON '07 - SharePointSitesExcluidos.json' criado com sucesso contendo todos os sites excluídos da coleção."

} catch {
    Write-Error "Erro ao obter informações da coleção de sites do SharePoint: $_"
}
