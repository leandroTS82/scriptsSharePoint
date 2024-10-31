# Iniciar o script
$scriptName = "05.1 - CriarCsvUsuarios"
Write-Output "$scriptName iniciado."

# Diretório para salvar os arquivos
$outputDirectory = ".\files\05json"

# Verifique se o diretório de saída existe, caso contrário, crie-o
if (-Not (Test-Path -Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory
}

# Caminho para o arquivo JSON
$jsonFilePath = Join-Path -Path $outputDirectory -ChildPath "05 - SharePointUsuarios_ComEmail.json"
# Caminho para o arquivo de texto
$siteNameFilePath = Join-Path -Path $outputDirectory -ChildPath "NomeDoSite.txt"

# Verifique se o arquivo JSON existe
if (-Not (Test-Path -Path $jsonFilePath)) {
    Write-Error "O arquivo JSON não foi encontrado: $jsonFilePath"
    exit
}

# Leia o conteúdo do arquivo JSON
$jsonContent = Get-Content -Path $jsonFilePath -Raw | ConvertFrom-Json

# Verifique se o arquivo de texto existe
if (-Not (Test-Path -Path $siteNameFilePath)) {
    Write-Error "O arquivo de texto 'NomeDoSite.txt' não foi encontrado: $siteNameFilePath"
    exit
}

# Leia o nome do site do arquivo de texto
$siteName = Get-Content -Path $siteNameFilePath -Raw
$siteName = $siteName.Trim()  # Remove espaços em branco

# Caminho completo para o arquivo CSV, incluindo o texto do nome do site
$csvFilePath = Join-Path -Path $outputDirectory -ChildPath ("05_usuarios_" + $siteName + ".csv")

# Tente criar o arquivo CSV
try {
    # Crie um array para armazenar os dados dos usuários
    $userData = @()

    # Preencha o array com os dados do JSON
    foreach ($user in $jsonContent) {
        $userData += [PSCustomObject]@{
            Nome  = $user.Nome
            Email = $user.Email
            Login = $user.Login
        }
    }

    # Exporta os dados para um arquivo CSV
    $userData | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
    Write-Output "Arquivo CSV '$csvFilePath' criado com sucesso."
} catch {
    Write-Error "Erro ao gerar o CSV: $_"
}

# Fim do script
Write-Output "$scriptName finalizado."
