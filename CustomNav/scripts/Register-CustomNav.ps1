# Register-CustomNav.ps1
# Directly registers the CustomNav Application Customizer on a SharePoint site using REST API via Azure CLI.
# Must run "az login" first.
#
# Usage:
#   .\Register-CustomNav.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite"

param([Parameter(Mandatory = $true)][string]$SiteUrl)

$componentId   = "3b6f2fb3-8cd0-4f52-9d2e-0d2a4f24c8d1"
$actionName    = "CustomNav"
$location      = "ClientSideExtension.ApplicationCustomizer"

# Verify az CLI is logged in
Write-Host "Verifying Azure CLI login..." -ForegroundColor Cyan
$account = az account show 2>$null | ConvertFrom-Json
if (-not $account) {
    Write-Host "Not logged in. Run 'az login' first." -ForegroundColor Red
    exit 1
}
Write-Host "Logged in as: $($account.user.name)" -ForegroundColor Green

# Fetch token for SharePoint using the correct scope
Write-Host "Fetching SharePoint access token..." -ForegroundColor Cyan
try {
    $tokenJson = az account get-access-token --scope "https://bfsaulco.sharepoint.com/.default" 2>$null | ConvertFrom-Json
    $token = $tokenJson.accessToken
    if (-not $token) { throw "Empty token" }

    # Decode JWT payload to verify audience
    $parts   = $token.Split('.')
    $payload = $parts[1].Replace('-','+').Replace('_','/')
    switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } }
    $claims  = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
    Write-Host "Token audience (aud): $($claims.aud)" -ForegroundColor DarkGray
    Write-Host "Token subject  (upn): $($claims.upn)" -ForegroundColor DarkGray
} catch {
    Write-Host "Error: Could not get SharePoint token. Ensure your account has SharePoint access." -ForegroundColor Red
    exit 1
}

$headers = @{
    "Authorization" = "Bearer $token"
    "Accept"        = "application/json;odata=nometadata"
    "Content-Type"  = "application/json"
}

# Test connectivity first
Write-Host "Testing SharePoint connectivity..." -ForegroundColor Cyan
try {
    $webInfo = Invoke-RestMethod -Method Get -Uri "$SiteUrl/_api/web?`$select=Title" -Headers $headers -ErrorAction Stop
    Write-Host "Connected to site: $($webInfo.Title)" -ForegroundColor Green
} catch {
    $statusCode = [int]$_.Exception.Response.StatusCode
    Write-Host "Cannot reach $SiteUrl (HTTP $statusCode)" -ForegroundColor Red
    
    # Try root site to check general SP access
    try {
        $rootInfo = Invoke-RestMethod -Method Get -Uri "https://bfsaulco.sharepoint.com/_api/web?`$select=Title" -Headers $headers -ErrorAction Stop
        Write-Host "Root site accessible: $($rootInfo.Title)" -ForegroundColor Yellow
        Write-Host "The account has SharePoint access but not to this specific site." -ForegroundColor Yellow
    } catch {
        $rootStatus = [int]$_.Exception.Response.StatusCode
        Write-Host "Root SharePoint also failed (HTTP $rootStatus) - token audience is likely wrong." -ForegroundColor Red
        Write-Host "Try: az logout, then az login, then re-run this script." -ForegroundColor Yellow
        exit 1
    }
    exit 1
}

$restUrl = "$SiteUrl/_api/web/UserCustomActions"

# Remove existing custom action if it exists
Write-Host "Checking for existing custom actions..." -ForegroundColor Yellow
try {
    $actionsResponse = Invoke-RestMethod -Method Get -Uri "$restUrl?`$filter=Name eq '$actionName'" `
        -Headers $headers -ErrorAction SilentlyContinue
    if ($actionsResponse.value -and $actionsResponse.value.Count -gt 0) {
        foreach ($action in $actionsResponse.value) {
            Write-Host "Removing existing action: $($action.Id)" -ForegroundColor Yellow
            Invoke-RestMethod -Method Delete -Uri "$restUrl('$($action.Id)')" -Headers $headers | Out-Null
        }
    }
} catch {
    # No existing actions, continue
}

# Register the custom action
Write-Host "Registering CustomNav..." -ForegroundColor Cyan
$body = @{
    Title                         = $actionName
    Name                          = $actionName
    Location                      = $location
    ClientSideComponentId         = $componentId
    ClientSideComponentProperties = "{}"
} | ConvertTo-Json -Compress

try {
    Invoke-RestMethod -Method Post -Uri $restUrl -Headers $headers -Body $body -ErrorAction Stop | Out-Null
    Write-Host "Success! Refresh $SiteUrl to verify the nav bar appears." -ForegroundColor Green
} catch {
    $status = $_.Exception.Response.StatusCode
    $msg    = $_.Exception.Message
    Write-Host "Error ($status): $msg" -ForegroundColor Red
    try {
        $stream = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream)
        Write-Host "Details: $($reader.ReadToEnd())" -ForegroundColor Yellow
    } catch {}
    exit 1
}
