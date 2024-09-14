# Importar o módulo necessário
Import-Module PnP.PowerShell

# Parâmetros de Conexão
$siteUrl = "https://ltsconsultoria.sharepoint.com"

# Conectar ao SharePoint Online usando autenticação interativa
try {
    Connect-PnPOnline -Url $siteUrl -UseWebLogin
    Write-Host "Conexão com SharePoint Online realizada com sucesso."

    # Obter as listas do SharePoint
    $lists = Get-PnPList
    if ($lists) {
        Write-Host "Listas obtidas com sucesso:"
        $lists | ForEach-Object {
            Write-Host $_.Title
        }
    } else {
        Write-Host "Nenhuma lista encontrada."
    }

} catch {
    Write-Error "Erro ao obter listas do SharePoint: $_"
}

# Desconectar ao final
Disconnect-PnPOnline
