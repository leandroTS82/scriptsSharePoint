# Conecte-se ao site do SharePoint Online
Connect-PnPOnline -Url "https://butterflygrowth.sharepoint.com/sites/Hopere" -UseWebLogin

# Diretório de saída
$outputDirectory = ".\files\alerts_json"

# Crie o diretório, se não existir
if (!(Test-Path -Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory
}

# Lista de e-mails dos usuários a serem verificados
$users = @(
    "leandrots82@gmail.com"
    # Adicione mais conforme necessário
)

# Lista para armazenar os alertas
$alertsList = @()

foreach ($userEmail in $users) {
    try {
        # Obtenha alertas do usuário
        $alerts = Get-PnPAlert -User $userEmail

        foreach ($alert in $alerts) {
            $alertObj = [PSCustomObject]@{
                Usuario             = $userEmail
                TituloDoAlerta      = $alert.Title
                URLMonitorada       = $alert.ListUrl
                TipoDeAlteracao     = $alert.ChangeType
                Frequencia          = $alert.AlertFrequency
                Tipo                = $alert.AlertType
                CriadoEm            = $alert.CreationDate
            }
            $alertsList += $alertObj
        }
    } catch {
        Write-Warning "Não foi possível recuperar alertas para $userEmail: $_"
    }
}

# Exporta o resultado para JSON
$jsonFilePath = Join-Path -Path $outputDirectory -ChildPath "07 - SharePointAlertsPorUsuario.json"
$alertsList | ConvertTo-Json -Depth 3 | Out-File -FilePath $jsonFilePath -Encoding UTF8

Write-Output "Arquivo JSON criado com sucesso contendo os alertas por usuário."
