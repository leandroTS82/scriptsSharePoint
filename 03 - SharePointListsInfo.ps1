# Importar o módulo necessário
Import-Module PnP.PowerShell

# Parâmetros de Conexão
$sitePai = "https://ltsconsultoria.sharepoint.com"
$subSite = "" #inicie sempre com a /,  se não tiver subsite deixe vazio ""
$siteUrl = $sitePai + $subSite

$outputDir = "./files"
# Verificar se o diretório de saída existe, caso contrário, criá-lo
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

$outputFile = "$outputDir/03 - $($subSite.TrimStart('/').Replace('/', '_'))SPListsInfo.json"
$ignoredFile = "$outputDir/03 - $($subSite.TrimStart('/').Replace('/', '_'))IgnoredItems.json"

# Arrays de exclusão de listas e colunas
$excludedLists = @("appdata", "appfiles", "SharePointHomeOrgLinks", "TaxonomyHiddenList", 
                    "Galeria de Temas", "Lista de Informações sobre o Usuário", 
                    "Galeria de Web Parts", "Galeria de Páginas Mestras", "Galeria de Soluções", "Aparências compostas", "Formulários Convertidos", "Web Template Extensions")  # Títulos ou InternalNames das listas a serem ignoradas
$excludedColumns = @("_Emoji", "_ColorHex", "_ColorTag", "ComplianceAssetId", "_HasCopyDestinations", "_CopySource", "owshiddenversion")  # Títulos ou InternalNames das colunas a serem ignoradas

# Arrays para armazenar listas e colunas ignoradas
$ignoredListsFound = @()
$ignoredListsNotFound = @()
$ignoredColumnsFound = @()
$ignoredColumnsNotFound = @()

# Função para imprimir mensagens em cores
function Write-Feedback {
    param (
        [string]$Message,
        [string]$Color = "White"
    )
    $originalColor = $Host.UI.RawUI.ForegroundColor
    $Host.UI.RawUI.ForegroundColor = [ConsoleColor]::Parse([ConsoleColor], $Color, $false)
    Write-Host $Message
    $Host.UI.RawUI.ForegroundColor = $originalColor
}

# Função para obter detalhes de uma lista de Lookup
function Get-LookupListDetails {
    param (
        [string]$lookupListId
    )
    
    if (-not [string]::IsNullOrEmpty($lookupListId)) {
        try {
            $lookupList = Get-PnPList -Identity $lookupListId
            return [PSCustomObject]@{
                Title = $lookupList.Title
                InternalName = $lookupList.InternalName
            }
        } catch {
            Write-Feedback "Erro ao obter detalhes da lista de Lookup com ID $lookupListId : $_" 
            return [PSCustomObject]@{
                Title = "Erro"
                InternalName = "Erro"
            }
        }
    } else {
        return [PSCustomObject]@{
            Title = "N/A"
            InternalName = "N/A"
        }
    }
}

# Conectar ao SharePoint Online usando autenticação interativa
try {
    Connect-PnPOnline -Url $siteUrl -UseWebLogin
    Write-Feedback "Conexão com SharePoint Online realizada com sucesso." -Color "Green"

    # Obter as listas do SharePoint
    $lists = Get-PnPList
    $listInfo = @()

    if ($lists) {
        Write-Feedback "Listas obtidas com sucesso." -Color "Green"

        foreach ($list in $lists) {
            # Verificar se a lista está no array de exclusão
            if ($excludedLists -contains $list.Title -or $excludedLists -contains $list.InternalName) {
                $ignoredListsFound += $list.Title
                Write-Feedback "Lista '$($list.Title)' ignorada." -Color "Yellow"
                continue
            }

            try {
                # Obter informações da lista
                $listTitle = $list.Title
                $listInternalName = $list.InternalName
                $listUrl = $list.Url
                $listAuthor = $list.Author -join ', '

                # Obter contagem de itens
                $listItemCount = (Get-PnPListItem -List $listTitle -PageSize 1 -ScriptBlock { $_ }).Count

                # Obter colunas
                $listColumns = Get-PnPField -List $listTitle
                $columnsInfo = @()

                foreach ($column in $listColumns) {
                    # Verificar se a coluna está no array de exclusão
                    if ($excludedColumns -contains $column.Title -or $excludedColumns -contains $column.InternalName) {
                        $ignoredColumnsFound += $column.Title
                        Write-Feedback "Coluna '$($column.Title)' ignorada." -Color "Yellow"
                        continue
                    }

                    $columnInfo = [PSCustomObject]@{
                        Title = $column.Title
                        InternalName = $column.InternalName
                        Type = $column.TypeAsString
                    }

                    # Adicionar informação de lista de origem para campos de Lookup
                    if ($column.TypeAsString -eq "Lookup" -and $column.LookupList) {
                        $lookupListDetails = Get-LookupListDetails -lookupListId $column.LookupList
                        $columnInfo | Add-Member -MemberType NoteProperty -Name "LookupListName" -Value $lookupListDetails.Title
                        $columnInfo | Add-Member -MemberType NoteProperty -Name "LookupListInternalName" -Value $lookupListDetails.InternalName
                    } else {
                        $columnInfo | Add-Member -MemberType NoteProperty -Name "LookupListName" -Value "N/A"
                        $columnInfo | Add-Member -MemberType NoteProperty -Name "LookupListInternalName" -Value "N/A"
                    }

                    $columnsInfo += $columnInfo
                }

                $listInfo += [PSCustomObject]@{
                    Title = $listTitle
                    InternalName = $listInternalName
                    Url = $listUrl
                    Author = $listAuthor
                    ItemCount = $listItemCount
                    Columns = $columnsInfo
                }
            } catch {
                Write-Feedback "Erro ao processar a lista $($list.Title): $_" -Color "Red"
            }
        }

        # Salvar as informações em um arquivo JSON
        $listInfo | ConvertTo-Json -Depth 5 | Out-File -FilePath $outputFile
        Write-Feedback "Informações das listas salvas em $outputFile" -Color "Green"

        # Verificar listas e colunas não encontradas
        $ignoredListsNotFound = $excludedLists | Where-Object { $_ -notin $ignoredListsFound }
        $ignoredColumnsNotFound = $excludedColumns | Where-Object { $_ -notin $ignoredColumnsFound }

        # Gerar arquivo JSON para listas e colunas ignoradas
        $ignoredInfo = [PSCustomObject]@{
            IgnoredListsFound = $ignoredListsFound
            IgnoredListsNotFound = $ignoredListsNotFound
            IgnoredColumnsFound = $ignoredColumnsFound
            IgnoredColumnsNotFound = $ignoredColumnsNotFound
        }
        $ignoredInfo | ConvertTo-Json -Depth 5 | Out-File -FilePath $ignoredFile
        Write-Feedback "Informações de listas e colunas ignoradas salvas em $ignoredFile" -Color "Green"
    } else {
        Write-Feedback "Nenhuma lista encontrada." -Color "Yellow"
    }
} catch {
    Write-Feedback "Erro ao obter listas do SharePoint: $_" -Color "Red"
} finally {
    # Desconectar ao final
    Disconnect-PnPOnline
}
