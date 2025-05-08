# Importar o módulo necessário
Import-Module PnP.PowerShell

# Parâmetros de Conexão
$siteUrl = "https://ltsconsultoria.sharepoint.com"

# Conectar ao SharePoint Online usando autenticação interativa
try {
    Connect-PnPOnline -Url $siteUrl -UseWebLogin
    Write-Host "Conexão com SharePoint Online realizada com sucesso."
} catch {
    Write-Error "Erro ao conectar ao SharePoint: $_"
}

# Desconectar ao final
Disconnect-PnPOnline
