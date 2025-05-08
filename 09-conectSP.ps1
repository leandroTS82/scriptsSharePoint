# 09-conectSP.ps1

$siteUrl = "https://ltsconsultoria.sharepoint.com/sites/PenielFloridaMirin"
$clientId = "ZNV8Q~KrEurjkVcE7LAC2HaeSamZY-qoDgfwbaB9"
$clientSecret = "0db3cf32-6ce0-4ef2-8ef3-f525da3aaff3"
$tenant = "ltsconsultoria.onmicrosoft.com"

Connect-PnPOnline -Url $siteUrl `
    -ClientId $clientId `
    -ClientSecret $clientSecret `
    -Tenant $tenant `
    -ClientSecretCredential
