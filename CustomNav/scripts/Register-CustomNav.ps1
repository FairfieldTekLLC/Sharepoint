# Register-CustomNav.ps1
# Directly registers the CustomNav Application Customizer on a SharePoint site.
# This bypasses app installation and feature activation.
#
# Prerequisites:
#   Install-Module PnP.PowerShell -Scope CurrentUser -Force
#
# Usage:
#   .\Register-CustomNav.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite"

param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl
)

$componentId   = "3b6f2fb3-8cd0-4f52-9d2e-0d2a4f24c8d1"
$actionName    = "CustomNav"
$location      = "ClientSideExtension.ApplicationCustomizer"

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
# Opens a browser login prompt — sign in with your SharePoint Admin account.
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

# Remove any stale registrations with the same name first
$existing = Get-PnPCustomAction -Scope Site | Where-Object { $_.Name -eq $actionName }
if ($existing) {
    Write-Host "Removing existing '$actionName' custom action ..." -ForegroundColor Yellow
    $existing | ForEach-Object { Remove-PnPCustomAction -Identity $_.Id -Scope Site -Force }
}

Write-Host "Registering CustomNav Application Customizer ..." -ForegroundColor Cyan
Add-PnPCustomAction `
    -Name       $actionName `
    -Title      $actionName `
    -Location   $location `
    -ClientSideComponentId         $componentId `
    -ClientSideComponentProperties "{}" `
    -Scope      Site

Write-Host "Done. Refresh the site to verify the nav bar appears." -ForegroundColor Green
