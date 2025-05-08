# Configurações
$siteUrl = "https://butterflygrowth.sharepoint.com/sites/SoApe"
$listName = "2025"
$userEmail = "leandrots82@gmail.com"

# Conectar ao SharePoint Online
Connect-PnPOnline -Url $siteUrl -UseWebLogin

try {
    # Criar alerta com os parâmetros suportados
    Add-PnPAlert -User $userEmail -List $listName -ChangeType All

    Write-Host "✅ Alerta criado com sucesso para o usuário $userEmail na lista '$listName'." -ForegroundColor Green
}
catch {
    Write-Host "❌ Erro ao criar o alerta: $_" -ForegroundColor Red
}
