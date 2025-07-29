# Importar módulos necessários
Import-Module PnP.PowerShell -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

# Caminhos dos arquivos
$sitesFilePath = ".\Config\Sites.json"
$ignoreFilePath = ".\Config\IgnoreColumns.json"
$excelFilePath = ".\Excel\ChannelContacts.xlsx"
$schemaFilePath = ".\Schemas\ChannelContacts_SchemaListColumns.json"

# Etapa 1 - Seleção do site
try {
    $sites = Get-Content $sitesFilePath | ConvertFrom-Json
    if (-not $sites) {
        Write-Error "Nenhum site encontrado no arquivo Sites.json"
        exit
    }

    Write-Host "`nSelecione o site SharePoint para importar os dados:`n" -ForegroundColor Cyan
    for ($i = 0; $i -lt $sites.Count; $i++) {
        Write-Host "[$i] $($sites[$i].name)"
    }

    $selectedIndex = Read-Host "`nDigite o número correspondente"
    if ($selectedIndex -notmatch '^\d+$' -or $selectedIndex -ge $sites.Count) {
        Write-Error "Seleção inválida. Encerrando script."
        exit
    }

    $siteUrl = $sites[$selectedIndex].url
    Write-Host "Site selecionado: $($sites[$selectedIndex].name) - $siteUrl" -ForegroundColor Green
} catch {
    Write-Error "Erro ao carregar ou interpretar o arquivo Sites.json: $_"
    exit
}

# Etapa 2 - Conectar ao site
try {
    Connect-PnPOnline -Url $siteUrl -UseWebLogin
    Write-Host "Conectado com sucesso ao SharePoint em $siteUrl" -ForegroundColor Green
} catch {
    Write-Error "Erro ao conectar ao SharePoint: $_"
    exit
}

# Etapa 3 - Carregar listas do site
try {
    $lists = Get-PnPList | Where-Object { -not $_.Hidden -and $_.BaseTemplate -eq 100 }
    if (-not $lists) {
        Write-Error "Nenhuma lista visível do tipo padrão encontrada no site."
        Disconnect-PnPOnline
        exit
    }

    Write-Host "`nSelecione a lista para importar os dados:`n" -ForegroundColor Cyan
    for ($i = 0; $i -lt $lists.Count; $i++) {
        Write-Host "[$i] $($lists[$i].Title)"
    }

    $listIndex = Read-Host "`nDigite o número da lista desejada"
    if (-not [int]::TryParse($listIndex, [ref]$null) -or [int]$listIndex -ge $lists.Count) {
        Write-Error "Seleção de lista inválida. Encerrando script."
        Disconnect-PnPOnline
        exit
    }
    $listIndex = [int]$listIndex
    $listId = $lists[$listIndex].Id
    $listName = $lists[$listIndex].Title
    Write-Host "Lista selecionada: $listName - ID: $listId" -ForegroundColor Green
} catch {
    Write-Error "Erro ao carregar listas do site: $_"
    Disconnect-PnPOnline
    exit
}

# Etapa 4 - Carregar arquivos auxiliares
try {
    $ignoredColumns = Get-Content $ignoreFilePath | ConvertFrom-Json
    Write-Host "Arquivo IgnoreColumns.json carregado com sucesso." -ForegroundColor Green

    $schema = Get-Content $schemaFilePath | ConvertFrom-Json
    $mappedColumns = $schema.Colunas | Where-Object {
        $_.InternalName -and $_.Title -and ($_.InternalName -notin $ignoredColumns)
    }
    Write-Host "SchemaListColumns.json carregado. Total de colunas consideradas: $($mappedColumns.Count)" -ForegroundColor Green
} catch {
    Write-Error "Erro ao carregar arquivos JSON auxiliares: $_"
    Disconnect-PnPOnline
    exit
}

# Etapa 5 - Ler planilha Excel
try {
    $sheetData = Import-Excel -Path $excelFilePath
    Write-Host "Planilha Excel carregada com sucesso. Total de linhas: $($sheetData.Count)" -ForegroundColor Green
} catch {
    Write-Error "Erro ao carregar a planilha Excel: $_"
    Disconnect-PnPOnline
    exit
}

# Etapa 6 - Inserir dados na lista
$itemCount = 0
foreach ($row in $sheetData) {
    $newItem = @{}

    foreach ($column in $mappedColumns) {
        $excelValue = $row.($column.Title)

        switch ($column.FieldTypeKind) {
            4 {
                if ($excelValue -and ($excelValue -is [datetime] -or ([datetime]::TryParse($excelValue, [ref]$null)))) {
                    $newItem[$column.InternalName] = [datetime]$excelValue
                } else {
                    Write-Warning "Coluna '$($column.Title)' esperava tipo [datetime], mas recebeu: '$excelValue'"
                    $newItem[$column.InternalName] = $null
                }
            }
            6 { $newItem[$column.InternalName] = if ($excelValue) { "$excelValue" } else { $null } }
            9 {
                $newItem[$column.InternalName] = if ($excelValue -match '^\d+([.,]\d+)?$') { [double]$excelValue } else {
                    Write-Warning "Coluna '$($column.Title)' esperava tipo [number], mas recebeu: '$excelValue'"
                    $null
                }
            }
            10 {
                $newItem[$column.InternalName] = if ($excelValue -match '^\d+([.,]\d+)?$') { [decimal]$excelValue } else {
                    Write-Warning "Coluna '$($column.Title)' esperava tipo [currency], mas recebeu: '$excelValue'"
                    $null
                }
            }
            2 { $newItem[$column.InternalName] = if ($excelValue) { "$excelValue" } else { $null } }
            3 { $newItem[$column.InternalName] = if ($excelValue) { "$excelValue" } else { $null } }
            default {
                Write-Host "Campo ignorado (não mapeado ou não suportado): $($column.InternalName)" -ForegroundColor Yellow
            }
        }
    }

    try {
        Add-PnPListItem -List $listId -Values $newItem
        $itemCount++
        Write-Host "Item $itemCount importado com sucesso." -ForegroundColor Cyan
    } catch {
        Write-Warning "Erro ao importar o item da linha $($itemCount + 1): $_"
    }
}

Write-Host "Importação finalizada. Total de itens importados: $itemCount" -ForegroundColor Green

# Desconectar
Disconnect-PnPOnline
