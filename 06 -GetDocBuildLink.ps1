$SPWeb = "https://butterflygrowth.sharepoint.com"
$SPSite ="sites/Agger_"
# Conecte-se ao SharePoint Online usando credenciais
Connect-PnPOnline -Url "$SPWeb/$SPSite" -UseWebLogin

# Defina parte do nome do documento que deseja procurar
$documentName = ""

# Defina o diretório e o caminho do arquivo JSON
$directoryPath = "./files/doc_json"
$jsonFilePath = "$directoryPath/documents_with_links.json"

# Verifica se o diretório existe, caso contrário, cria
if (-not (Test-Path -Path $directoryPath)) {
    New-Item -ItemType Directory -Force -Path $directoryPath
}

# Cria uma lista para armazenar os resultados
$documentList = @()

# Obtém todas as bibliotecas de documentos do site
$lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }

# Defina a URL base do site manualmente para evitar duplicação

# Loop em cada biblioteca de documentos
foreach ($list in $lists) {
    Write-Host "Procurando na biblioteca: $($list.Title)"
    
    # Obtém todos os arquivos da biblioteca
    $allItems = Get-PnPListItem -List $list -PageSize 100

    foreach ($item in $allItems) {
        $fileName = $item.FieldValues["FileLeafRef"]
        Write-Host "Documento encontrado: $fileName"

        # Verifica se o arquivo contém o nome especificado (ou se $documentName está vazio)
        if ($documentName -eq "" -or $fileName -like "*$documentName*") {
            Write-Host "Documento correspondente: $fileName"
            $fileUrl = $item.FieldValues["FileRef"]  # Obtém a URL relativa do documento
            $fullUrl = "$SPWeb$fileUrl"  # Combina a URL base com a URL relativa para criar a URL completa

            # Inicializa variáveis para controle de erro
            $shareLink = $null
            $externalAccessError = $false

            # Tenta gerar o link de compartilhamento externo
            try {
                # Verifica se o cmdlet está disponível
                if (Get-Command -Name Get-PnPSharingLink -ErrorAction SilentlyContinue) {
                    $shareLink = Get-PnPSharingLink -Url $fullUrl -SharingType View -AllowExternalSharing $true
                } else {
                    Write-Host "Get-PnPSharingLink não está disponível. Gerando link de acesso normal."
                }
            } catch {
                Write-Host "Erro ao tentar gerar o link de compartilhamento: $_"
                $externalAccessError = $true
            }

            # Se o link de compartilhamento não foi gerado, retorna a URL do documento
            if (-not $shareLink) {
                $externalAccessError = $true
                Write-Host "Não foi possível gerar o link de compartilhamento para $fileName. Usando a URL do documento."
                $shareLink = @{ Link = $fullUrl }  # Define o link como a URL completa do documento
            }

            # Adiciona os detalhes do documento à lista
            $documentList += @{
                "FileName" = $fileName
                "FileUrl" = $fullUrl
                "SharingLink" = $shareLink.Link
                "ExternalAccessError" = $externalAccessError
            }
            Write-Host "Adicionado ao JSON: $fileName"
        }
    }
}

# Verifica se algum documento foi encontrado e salva no JSON
if ($documentList.Count -gt 0) {
    Write-Host "Gerando JSON com os documentos encontrados..."

    # Converte a lista em JSON e salva no arquivo
    $documentListJson = $documentList | ConvertTo-Json -Depth 5
    
    # Salva o JSON no arquivo
    Set-Content -Path $jsonFilePath -Value $documentListJson -Force
    
    Write-Host "Os links de compartilhamento foram salvos em: $jsonFilePath"
} else {
    Write-Host "Nenhum documento correspondente foi encontrado."
}

# Desconectar do SharePoint Online
Disconnect-PnPOnline
