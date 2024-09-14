# Conecte-se ao site do SharePoint Online
Connect-PnPOnline -Url "https://ltsconsultoria.sharepoint.com/sites/controlefinanceiro" -UseWebLogin

# Obtenha informações do site
try {
    # Obter informações do site
    $siteInfo = Get-PnPSite

    # Obter informações adicionais do site
    $webInfo = Get-PnPWeb

    # Crie um objeto para armazenar as informações do site
    $siteData = [PSCustomObject]@{
        TituloSite           = $siteInfo.Title
        URLSite              = $siteInfo.Url
        DescricaoSite        = $siteInfo.Description
        CriadoEm             = $siteInfo.Created
        UltimaModificacao    = $siteInfo.LastItemModifiedDate
        Idioma               = $siteInfo.Language
        TituloWeb            = $webInfo.Title
        URLWeb               = $webInfo.Url
        DescricaoWeb         = $webInfo.Description
        LogoURL              = $webInfo.LogoUrl
        VersaoWeb            = $webInfo.WebTemplate
    }

    # Converta o objeto para JSON e salve em um arquivo
    $siteData | ConvertTo-Json -Depth 3 | Out-File -FilePath "files/04 - SharePointSiteInfo.json" -Encoding UTF8

    Write-Output "Arquivo JSON '04 - SharePointSiteInfo.json' criado com sucesso."
}
catch {
    Write-Error "Erro ao obter informações do SharePoint: $_"
}
