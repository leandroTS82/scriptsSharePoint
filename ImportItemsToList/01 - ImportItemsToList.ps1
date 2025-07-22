# Importar módulos necessários
Import-Module PnP.PowerShell -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

# Caminhos fixos
$sitesFilePath = ".\Sites.json"
$schemaFilePath = ".\SchemaListColumns.json"
$ignoreFilePath = ".\IgnoreColumns.json"
$excelFilePath = ".\Excel\DiarioBordo.xlsx"

# Selecionar site interativamente
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

# Conectar ao SharePoint
try {
    Connect-PnPOnline -Url $siteUrl -UseWebLogin
    Write-Host "Conectado com sucesso ao SharePoint em $siteUrl" -ForegroundColor Green
} catch {
    Write-Error "Erro ao conectar ao SharePoint: $_"
    exit
}

# Carregar colunas a serem ignoradas
try {
    $ignoredColumns = Get-Content $ignoreFilePath | ConvertFrom-Json
    Write-Host "Arquivo de colunas ignoradas carregado com sucesso." -ForegroundColor Green
} catch {
    Write-Error "Erro ao carregar o arquivo IgnoreColumns.json: $_"
    Disconnect-PnPOnline
    exit
}

# Carregar schema com listId e colunas
try {
    $schema = Get-Content $schemaFilePath | ConvertFrom-Json
    $listId = $schema.listId
    $mappedColumns = $schema.Colunas | Where-Object {
        $_.InternalName -and $_.Title -and ($_.InternalName -notin $ignoredColumns)
    }
    Write-Host "Schema de colunas carregado. Total de colunas consideradas: $($mappedColumns.Count)" -ForegroundColor Green
} catch {
    Write-Error "Erro ao carregar o arquivo SchemaListColumns.json: $_"
    Disconnect-PnPOnline
    exit
}

# Verificar se a lista existe
$listExists = Get-PnPList -Identity $listId -ErrorAction SilentlyContinue
if (-not $listExists) {
    Write-Error "A lista com ID '$listId' não foi encontrada no site $siteUrl"
    Disconnect-PnPOnline
    exit
}

# Ler os dados da planilha Excel
try {
    $sheetData = Import-Excel -Path $excelFilePath
    Write-Host "Planilha Excel carregada com sucesso. Total de linhas: $($sheetData.Count)" -ForegroundColor Green
} catch {
    Write-Error "Erro ao carregar a planilha Excel: $_"
    Disconnect-PnPOnline
    exit
}

# Inserir os dados na lista do SharePoint
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
