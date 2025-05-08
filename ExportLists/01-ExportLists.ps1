# Definição das variáveis
$siteUrl = "https://butterflygrowth.sharepoint.com/sites/Agger_"  # Substitua pela URL do site
$listName = "Cronograma Inbound"                                  # Substitua pelo nome da lista
$templateFileName = "Cronograma Inbound-template.xml"             # Nome do arquivo template
$outputFolder = "./files"                                         # Pasta onde o arquivo será salvo

# Verificar se a pasta existe, caso contrário, criá-la
if (-Not (Test-Path -Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder
}

# Conectar ao SharePoint Online usando o modo interativo recomendado
Connect-PnPOnline -Url $siteUrl -Interactive

# Exportar a lista como um template
Get-PnPProvisioningTemplate -Out "$outputFolder/$templateFileName" -Handlers Lists -List $listName

Write-Host "Template da lista exportado para a pasta: $outputFolder"
