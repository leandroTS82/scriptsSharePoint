# Importar o módulo necessário
Import-Module PnP.PowerShell

# Parâmetros de Conexão
$siteUrl = "https://ltsconsultoria.sharepoint.com"
$outputFile = "./files/03 - SharePointListsInfo.json"

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
    } else {
        Write-Feedback "Nenhuma lista encontrada." -Color "Yellow"
    }
} catch {
    Write-Feedback "Erro ao obter listas do SharePoint: $_" -Color "Red"
} finally {
    # Desconectar ao final
    Disconnect-PnPOnline
}
