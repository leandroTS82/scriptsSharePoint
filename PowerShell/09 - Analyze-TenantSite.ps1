# CONFIGURAÇÃO
$siteUrl = "https://ltsconsultoria.sharepoint.com/sites/ControleFinanceiro"
$outputPath = ".\tenant-analysis.json"

# CONEXÃO
Connect-PnPOnline -Url $siteUrl -UseWebLogin

# OBJETO PARA COLETAR OS DADOS
$analysis = @{
    SiteUrl   = $siteUrl
    Timestamp = (Get-Date).ToString("s")
}

# OBTÉM INFORMAÇÕES DO SITE
$web = Get-PnPWeb
$site = Get-PnPTenantSite -Url $siteUrl

$analysis["Web"] = @{
    Url         = $web.Url
    Title       = $web.Title
    Description = $web.Description
}

$analysis["Site"] = @{
    Url                = $site.Url
    Status             = $site.Status
    SharingCapability  = $site.SharingCapability
    StorageQuota       = $site.StorageQuota
    StorageUsed        = $site.StorageUsage
    ResourceQuota      = $site.ResourceQuota
    AllowAppSideLoading = $site.AllowAppSideLoading
    LockState          = $site.LockState
    Template           = $site.Template
    Classification     = $site.Classification
    DisableCustomScript = $site.DisableCustomScript
}

# ADMINS
$admins = Get-PnPSiteCollectionAdmin | Select-Object LoginName, Email, IsSiteAdmin
$analysis["SiteAdmins"] = $admins

# USUÁRIOS EXTERNOS
$externalUsers = Get-PnPUser | Where-Object { $_.LoginName -like "*#ext#*" } | Select-Object Title, Email, LoginName
$analysis["ExternalUsers"] = $externalUsers

# GRUPOS E PERMISSÕES
$groups = @()
foreach ($group in Get-PnPGroup) {
    try {
        $members = Get-PnPGroupMembers -Identity $group | Select-Object -ExpandProperty Email
    } catch {
        $members = @("Erro ao obter membros ou cmdlet não disponível")
    }

    $groups += @{
        Group      = $group.Title
        Users      = $members
        Permissions = @("N/A")  # Pode ser ajustado caso implemente Get-PnPRoleAssignment
    }
}
$analysis["GroupsAndPermissions"] = $groups

# SOLUÇÕES SPFx INSTALADAS
try {
    $apps = Get-PnPApp | Select-Object Title, Id, InstalledVersion
    $analysis["SPFxClientSideSolutions"] = $apps
} catch {
    $analysis["SPFxClientSideSolutions"] = @("Erro ao obter soluções SPFx")
}

# SALVAR COMO JSON
$json = $analysis | ConvertTo-Json -Depth 6
$json | Out-File $outputPath -Encoding utf8

Write-Host "✅ Análise concluída. Arquivo gerado em: $outputPath" -ForegroundColor Green
