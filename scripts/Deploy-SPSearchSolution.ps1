<#
.SYNOPSIS
    Deploys the SP Search solution to a SharePoint Online site.

.DESCRIPTION
    End-to-end deployment script that:
    1. Ensures a site-level App Catalog exists (or uses a tenant-level one)
    2. Uploads and publishes the .sppkg package
    3. Installs the app on the target site
    4. Optionally applies the PnP provisioning template (lists, page, security, seed data)

.PARAMETER SiteUrl
    Target site URL where the app should be installed.
    Example: https://contoso.sharepoint.com/sites/search

.PARAMETER ClientId
    Azure AD app registration Client ID for PnP authentication.
    Uses browser-based interactive login with this Client ID.

.PARAMETER AppCatalogScope
    Where to deploy the .sppkg: "SiteLevel" (default) or "TenantLevel".
    SiteLevel creates/uses a site-level App Catalog on the target site.
    TenantLevel requires -AppCatalogUrl to be specified.

.PARAMETER AppCatalogUrl
    Tenant-level App Catalog URL. Required only when AppCatalogScope is "TenantLevel".
    Example: https://contoso.sharepoint.com/sites/appcatalog

.PARAMETER PackagePath
    Path to the .sppkg file. Defaults to sharepoint/solution/sp-search.sppkg.

.PARAMETER ProvisionSite
    Apply the PnP provisioning template after app install. This creates:
    - SP Search Admins security group
    - 3 hidden lists (SearchSavedQueries, SearchHistory, SearchCollections)
    - Search page with all web parts in optimal layout

.PARAMETER SkipInstall
    Skip installing the app on the target site (upload and publish only).

.EXAMPLE
    # Deploy app only (no site provisioning)
    .\scripts\Deploy-SPSearchSolution.ps1 `
      -SiteUrl "https://contoso.sharepoint.com/sites/search" `
      -ClientId "970bb320-0d49-4b4a-aa8f-c3f4b1e5928f"

.EXAMPLE
    # Full deploy: app + lists + page + security
    .\scripts\Deploy-SPSearchSolution.ps1 `
      -SiteUrl "https://contoso.sharepoint.com/sites/search" `
      -ClientId "970bb320-0d49-4b4a-aa8f-c3f4b1e5928f" `
      -ProvisionSite

.EXAMPLE
    # Tenant-level App Catalog
    .\scripts\Deploy-SPSearchSolution.ps1 `
      -SiteUrl "https://contoso.sharepoint.com/sites/search" `
      -ClientId "970bb320-0d49-4b4a-aa8f-c3f4b1e5928f" `
      -AppCatalogScope "TenantLevel" `
      -AppCatalogUrl "https://contoso.sharepoint.com/sites/appcatalog" `
      -ProvisionSite
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Target SharePoint site URL")]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true, HelpMessage = "Azure AD Client ID for PnP authentication")]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [ValidateSet("SiteLevel", "TenantLevel")]
    [string]$AppCatalogScope = "SiteLevel",

    [Parameter(Mandatory = $false)]
    [string]$AppCatalogUrl,

    [Parameter(Mandatory = $false)]
    [string]$PackagePath = (Join-Path $PSScriptRoot "..\sharepoint\solution\sp-search.sppkg"),

    [Parameter(Mandatory = $false)]
    [switch]$ProvisionSite,

    [Parameter(Mandatory = $false)]
    [switch]$SkipInstall
)

# ============================================================================
# Validate prerequisites
# ============================================================================
$ErrorActionPreference = "Stop"

$requiredModule = "PnP.PowerShell"
$module = Get-Module -ListAvailable -Name $requiredModule |
    Sort-Object Version -Descending |
    Select-Object -First 1

if (-not $module) {
    throw "PnP.PowerShell module not found. Install it with: Install-Module -Name PnP.PowerShell -Scope CurrentUser"
}

Import-Module $module.Path -Force -ErrorAction Stop

if (-not (Test-Path $PackagePath)) {
    throw "Package not found at: $PackagePath`nRun 'gulp bundle --ship && gulp package-solution --ship' first."
}

if ($AppCatalogScope -eq "TenantLevel" -and -not $AppCatalogUrl) {
    throw "AppCatalogUrl is required when AppCatalogScope is 'TenantLevel'."
}

$PackagePath = Resolve-Path $PackagePath
$packageSize = [math]::Round((Get-Item $PackagePath).Length / 1MB, 1)

# Provisioning template path
$templateDir = Join-Path $PSScriptRoot "..\provisioning"
$templatePath = Join-Path $templateDir "SiteTemplate.xml"
if ($ProvisionSite -and -not (Test-Path $templatePath)) {
    throw "Provisioning template not found at: $templatePath"
}

# ============================================================================
# XInclude resolver (pre-processes template before applying)
# ============================================================================
function Resolve-XIncludes {
    param(
        [string]$TemplatePath,
        [string]$OutputPath
    )

    $baseDir = Split-Path -Parent $TemplatePath
    $content = Get-Content -Path $TemplatePath -Raw

    $script:xiIncludeCount = 0
    $xiPattern = '<xi:include\s+href="([^"]+)"\s*/?\s*>'

    while ($content -match $xiPattern) {
        $content = [regex]::Replace($content, $xiPattern, {
            param($match)
            $href = $match.Groups[1].Value
            $includePath = Join-Path $baseDir $href

            if (Test-Path $includePath) {
                $includeContent = Get-Content -Path $includePath -Raw
                # Strip XML declarations from included files
                $includeContent = $includeContent -replace '<\?xml[^?]*\?>', ''
                $script:xiIncludeCount++
                return $includeContent.Trim()
            } else {
                Write-Warning "XInclude file not found: $includePath"
                return "<!-- MISSING: $href -->"
            }
        })
    }

    # Remove XInclude namespace declaration (no longer needed after resolution)
    $content = $content -replace '\s*xmlns:xi="http://www.w3.org/2001/XInclude"', ''

    Set-Content -Path $OutputPath -Value $content -Encoding UTF8
    return $script:xiIncludeCount
}

# ============================================================================
# Determine step count
# ============================================================================
$totalSteps = 4
if ($ProvisionSite) { $totalSteps++ }
$step = 0

try {
    Write-Host ""
    Write-Host "======================================================================" -ForegroundColor Cyan
    Write-Host " SP Search — Solution Deployment" -ForegroundColor Cyan
    Write-Host "======================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Site:      $SiteUrl" -ForegroundColor White
    Write-Host "  Catalog:   $AppCatalogScope" -ForegroundColor White
    Write-Host "  Package:   $PackagePath ($packageSize MB)" -ForegroundColor White
    Write-Host "  Client ID: $ClientId" -ForegroundColor White
    if ($ProvisionSite) {
        Write-Host "  Provision: 3 Lists + Page + Security (PnP template)" -ForegroundColor White
    }
    Write-Host ""

    # ─── Step 1: Connect to SharePoint ───────────────────────────
    $step++
    Write-Host "[$step/$totalSteps] Connecting to SharePoint..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
    Write-Host "  Connected successfully" -ForegroundColor Green
    Write-Host ""

    # ─── Step 2: Ensure App Catalog ──────────────────────────────
    $step++
    if ($AppCatalogScope -eq "SiteLevel") {
        Write-Host "[$step/$totalSteps] Ensuring site-level App Catalog..." -ForegroundColor Cyan

        try {
            Add-PnPSiteCollectionAppCatalog -Site $SiteUrl -ErrorAction Stop
            Write-Host "  Site-level App Catalog enabled" -ForegroundColor Green
        } catch {
            if ($_.Exception.Message -match "already exists" -or $_.Exception.Message -match "already been added" -or $_.Exception.Message -match "duplicate values") {
                Write-Host "  [EXISTS] Site-level App Catalog already enabled" -ForegroundColor Yellow
            } else {
                throw
            }
        }

        $catalogUrl = $SiteUrl
    } else {
        Write-Host "[$step/$totalSteps] Using tenant-level App Catalog..." -ForegroundColor Cyan
        Write-Host "  Catalog URL: $AppCatalogUrl" -ForegroundColor White
        $catalogUrl = $AppCatalogUrl

        # Reconnect to the tenant App Catalog site
        Connect-PnPOnline -Url $catalogUrl -ClientId $ClientId -Interactive
    }
    Write-Host ""

    # ─── Step 3: Upload and publish ──────────────────────────────
    $step++
    Write-Host "[$step/$totalSteps] Uploading and publishing solution package ($packageSize MB)..." -ForegroundColor Cyan
    Write-Host "  This may take a few minutes..." -ForegroundColor Yellow

    # Upload via .NET HttpClient with custom timeout (PnP Add-PnPApp has a hard 200s HttpClient timeout)
    $token = Get-PnPAccessToken -ResourceTypeName SharePoint -ErrorAction Stop
    $fileName = [System.IO.Path]::GetFileName($PackagePath)

    if ($AppCatalogScope -eq "SiteLevel") {
        $uploadUrl = "$catalogUrl/_api/web/sitecollectionappcatalog/Add(overwrite=true, url='$fileName')"
    } else {
        $uploadUrl = "$catalogUrl/_api/web/tenantappcatalog/Add(overwrite=true, url='$fileName')"
    }

    Write-Host "  Uploading via REST API (10 min timeout)..." -ForegroundColor Yellow

    $httpClient = [System.Net.Http.HttpClient]::new()
    $httpClient.Timeout = [TimeSpan]::FromMinutes(10)
    $httpClient.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $token)
    $httpClient.DefaultRequestHeaders.Accept.Add([System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new("application/json"))

    try {
        $fileContent = [System.Net.Http.ByteArrayContent]::new([System.IO.File]::ReadAllBytes($PackagePath))
        $fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("application/octet-stream")

        $response = $httpClient.PostAsync($uploadUrl, $fileContent).GetAwaiter().GetResult()

        if (-not $response.IsSuccessStatusCode) {
            $body = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            throw "Upload failed: $($response.StatusCode) — $body"
        }

        $resultJson = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $uploadResult = $resultJson | ConvertFrom-Json
        $appId = $uploadResult.UniqueId
        if (-not $appId) { $appId = $uploadResult.d.UniqueId }

        Write-Host "  Uploaded:   $fileName" -ForegroundColor Green
        Write-Host "  App ID:     $appId" -ForegroundColor Green
    } finally {
        $httpClient.Dispose()
    }

    # Deploy (publish) the app via REST API (PnP's Publish-PnPApp also hits 200s timeout)
    Write-Host "  Publishing..." -ForegroundColor Yellow

    $httpClient2 = [System.Net.Http.HttpClient]::new()
    $httpClient2.Timeout = [TimeSpan]::FromMinutes(10)
    $httpClient2.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $token)
    $httpClient2.DefaultRequestHeaders.Accept.Add([System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new("application/json"))

    try {
        if ($AppCatalogScope -eq "SiteLevel") {
            $deployUrl = "$catalogUrl/_api/web/sitecollectionappcatalog/AvailableApps/GetById('$appId')/Deploy"
        } else {
            $deployUrl = "$catalogUrl/_api/web/tenantappcatalog/AvailableApps/GetById('$appId')/Deploy"
        }

        $emptyContent = [System.Net.Http.StringContent]::new("{}", [System.Text.Encoding]::UTF8, "application/json")
        $deployResponse = $httpClient2.PostAsync($deployUrl, $emptyContent).GetAwaiter().GetResult()

        if (-not $deployResponse.IsSuccessStatusCode) {
            $deployBody = $deployResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            throw "Deploy failed: $($deployResponse.StatusCode) — $deployBody"
        }

        Write-Host "  Published:  Yes" -ForegroundColor Green
    } finally {
        $httpClient2.Dispose()
    }
    Write-Host ""

    # ─── Step 4: Install on target site ──────────────────────────
    $step++
    if (-not $SkipInstall) {
        Write-Host "[$step/$totalSteps] Installing app on target site..." -ForegroundColor Cyan

        # Reconnect to target site if we were on tenant catalog
        if ($AppCatalogScope -eq "TenantLevel") {
            Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
        }

        # Check if already installed
        $installedApp = Get-PnPApp -Identity $appId -Scope Site -ErrorAction SilentlyContinue
        if ($installedApp -and $installedApp.InstalledVersion) {
            Write-Host "  [UPDATE] App already installed (v$($installedApp.InstalledVersion)), updating..." -ForegroundColor Yellow
            Update-PnPApp -Identity $appId -Scope Site -ErrorAction Stop
        } else {
            Install-PnPApp -Identity $appId -Scope Site -Wait -ErrorAction Stop
        }

        Write-Host "  Installed on $SiteUrl" -ForegroundColor Green
    } else {
        Write-Host "[$step/$totalSteps] Skipping site install (--SkipInstall)" -ForegroundColor Yellow
    }
    Write-Host ""

    # ─── Step 5: Apply provisioning template (optional) ──────────
    if ($ProvisionSite) {
        $step++
        Write-Host "[$step/$totalSteps] Applying PnP provisioning template..." -ForegroundColor Cyan

        # Resolve XIncludes into a single template file
        $resolvedPath = Join-Path $templateDir "SiteTemplate.resolved.xml"
        Write-Host "  Resolving XIncludes..." -ForegroundColor Yellow
        $includeCount = Resolve-XIncludes -TemplatePath $templatePath -OutputPath $resolvedPath
        Write-Host "  Resolved $includeCount XInclude directives" -ForegroundColor Green

        # Ensure connected to target site
        try {
            $null = Get-PnPConnection -ErrorAction Stop
        } catch {
            Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
        }

        # Apply the provisioning template
        Write-Host "  Applying template (lists, page, security)..." -ForegroundColor Yellow
        Invoke-PnPSiteTemplate -Path $resolvedPath -ErrorAction Stop

        # Clean up resolved file
        Remove-Item $resolvedPath -Force -ErrorAction SilentlyContinue

        Write-Host "  [OK] Security: SP Search Admins group" -ForegroundColor Green
        Write-Host "  [OK] Lists: SearchSavedQueries, SearchHistory, SearchCollections" -ForegroundColor Green
        Write-Host "  [OK] Page: Search.aspx with all web parts" -ForegroundColor Green

        # Post-provisioning: Set ReadSecurity/WriteSecurity on SearchHistory
        # PnP provisioning schema doesn't support these attributes, so we set them via CSOM
        Write-Host "  Configuring SearchHistory item-level security..." -ForegroundColor Yellow
        try {
            $historyList = Get-PnPList -Identity "Lists/SearchHistory" -ErrorAction Stop
            if ($historyList) {
                # ReadSecurity=2: Users can only read their own items
                # WriteSecurity=2: Users can only edit their own items
                $historyList.ReadSecurity = 2
                $historyList.WriteSecurity = 2
                $historyList.Update()
                Invoke-PnPQuery -ErrorAction Stop
                Write-Host "  [OK] SearchHistory: ReadSecurity=2, WriteSecurity=2 (user-only access)" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Could not set SearchHistory item-level security: $($_.Exception.Message)"
            Write-Host "  Set manually: List Settings > Advanced > Read/Edit access = 'Only their own'" -ForegroundColor Yellow
        }
        Write-Host ""
    }

    # ─── Summary ─────────────────────────────────────────────────
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host " Deployment completed successfully!" -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Green
    Write-Host ""

    if ($ProvisionSite) {
        Write-Host "  Search page: $SiteUrl/SitePages/Search.aspx" -ForegroundColor White
        Write-Host ""
        Write-Host "  ┌──────────────────────────────────────────────────────┐"
        Write-Host "  │ Search Box (full width)                              │"
        Write-Host "  ├──────────────────────────────────────────────────────┤"
        Write-Host "  │ Verticals: All │ Documents │ Pages │ People │ Sites │"
        Write-Host "  ├────────────────────────────┬─────────────────────────┤"
        Write-Host "  │ Results (66%)              │ Filters (33%)          │"
        Write-Host "  └────────────────────────────┴─────────────────────────┘"
        Write-Host ""
        Write-Host "  Customize web parts by editing the page in SharePoint." -ForegroundColor Gray
    } else {
        Write-Host "  Next: Run with -ProvisionSite to create lists + search page:" -ForegroundColor Yellow
        Write-Host "    .\scripts\Deploy-SPSearchSolution.ps1 -SiteUrl `"$SiteUrl`" -ClientId `"$ClientId`" -ProvisionSite" -ForegroundColor Gray
    }
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "======================================================================" -ForegroundColor Red
    Write-Host " Deployment failed at step $step!" -ForegroundColor Red
    Write-Host "======================================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red

    if ($_.Exception.InnerException) {
        Write-Host "Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "Stack:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    Write-Host ""

    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  - Verify you are a Site Collection Admin on $SiteUrl" -ForegroundColor White
    Write-Host "  - Verify Client ID $ClientId has the required API permissions" -ForegroundColor White
    Write-Host "  - For site-level catalog issues, try tenant-level instead" -ForegroundColor White
    Write-Host "  - Re-run this script — it is safe to re-run (idempotent)" -ForegroundColor White
    Write-Host ""
    exit 1

} finally {
    try {
        $null = Get-PnPConnection -ErrorAction Stop
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
    } catch {
        # Already disconnected
    }
}
