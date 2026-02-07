<#
.SYNOPSIS
    One-shot setup: deploy .sppkg + provision hidden lists + create search page with configured web parts.

.DESCRIPTION
    Combines Deploy-SPSearchSolution, Provision-SPSearchLists, and Provision-SPSearchPage
    into a single script with one auth prompt. Run this after 'gulp bundle --ship && gulp package-solution --ship'.

.PARAMETER SiteUrl
    Target SharePoint site URL.

.PARAMETER ClientId
    Azure AD app registration Client ID for PnP authentication.

.EXAMPLE
    .\scripts\Setup-SPSearchSite.ps1 `
        -SiteUrl "https://pixelboy.sharepoint.com/sites/SPSearch" `
        -ClientId "970bb320-0d49-4b4a-aa8f-c3f4b1e5928f"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$PackagePath = (Join-Path $PSScriptRoot "..\sharepoint\solution\sp-search.sppkg"),

    [Parameter(Mandatory = $false)]
    [string]$PageName = "Search",

    [Parameter(Mandatory = $false)]
    [string]$SearchContextId = "default"
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell -ErrorAction Stop

# Web Part Component IDs (from manifest files)
$WP_SEARCH_BOX       = "13a82dbe-2c57-4e20-bfe8-ec4de5776191"
$WP_SEARCH_RESULTS   = "1836671c-a710-45b4-9a83-55c65344a3d5"
$WP_SEARCH_FILTERS   = "2eb68250-879f-45a8-af9b-9fc3e97b2050"
$WP_SEARCH_VERTICALS = "d0481c49-49f9-4219-90fe-be8338051f58"
$WP_SEARCH_MANAGER   = "46308c1c-af6b-43c5-98b7-2d39082498cb"

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " SP Search — Full Site Setup" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Site:    $SiteUrl"
Write-Host "  Package: $PackagePath"
Write-Host ""

# ═══════════════════════════════════════════════════════════════════════
# PHASE 1: Connect
# ═══════════════════════════════════════════════════════════════════════
Write-Host "[1/5] Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
Write-Host "  Connected" -ForegroundColor Green

try {
    # ═══════════════════════════════════════════════════════════════════
    # PHASE 2: Deploy .sppkg
    # ═══════════════════════════════════════════════════════════════════
    Write-Host ""
    Write-Host "[2/5] Deploying solution package..." -ForegroundColor Cyan

    if (-not (Test-Path $PackagePath)) {
        throw "Package not found: $PackagePath — run 'gulp bundle --ship && gulp package-solution --ship' first"
    }

    $PackagePath = Resolve-Path $PackagePath
    $packageSize = [math]::Round((Get-Item $PackagePath).Length / 1MB, 1)
    Write-Host "  Package: $packageSize MB" -ForegroundColor Gray

    # Ensure site-level app catalog
    try {
        Add-PnPSiteCollectionAppCatalog -Site $SiteUrl -ErrorAction Stop
        Write-Host "  Site-level App Catalog enabled" -ForegroundColor Green
    } catch {
        if ($_.Exception.Message -match "already exists|already been added|duplicate") {
            Write-Host "  App Catalog already enabled" -ForegroundColor Yellow
        } else { throw }
    }

    # Wait for site-level app catalog to provision (can take 30-90 seconds)
    $token = Get-PnPAccessToken -ResourceTypeName SharePoint -ErrorAction Stop
    $catalogCheckUrl = "$SiteUrl/_api/web/sitecollectionappcatalog"
    $maxWait = 120
    $waited = 0
    $catalogReady = $false

    Write-Host "  Waiting for App Catalog to be ready..." -ForegroundColor Yellow -NoNewline
    while ($waited -lt $maxWait) {
        try {
            $checkClient = [System.Net.Http.HttpClient]::new()
            $checkClient.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $token)
            $checkClient.DefaultRequestHeaders.Accept.Add([System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new("application/json"))
            $checkResp = $checkClient.GetAsync($catalogCheckUrl).GetAwaiter().GetResult()
            $checkClient.Dispose()
            if ($checkResp.IsSuccessStatusCode) {
                $catalogReady = $true
                break
            }
        } catch { }
        Start-Sleep -Seconds 5
        $waited += 5
        Write-Host "." -NoNewline
    }
    Write-Host ""

    if (-not $catalogReady) {
        throw "App Catalog not ready after $maxWait seconds. Please try again in a few minutes."
    }
    Write-Host "  App Catalog ready" -ForegroundColor Green

    # Upload via REST API (bypasses PnP 200s timeout)
    $fileName = [System.IO.Path]::GetFileName($PackagePath)
    $uploadUrl = "$SiteUrl/_api/web/sitecollectionappcatalog/Add(overwrite=true, url='$fileName')"

    Write-Host "  Uploading..." -ForegroundColor Yellow
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
        Write-Host "  Uploaded (App ID: $appId)" -ForegroundColor Green
    } finally { $httpClient.Dispose() }

    # Deploy (publish)
    Write-Host "  Publishing..." -ForegroundColor Yellow
    $httpClient2 = [System.Net.Http.HttpClient]::new()
    $httpClient2.Timeout = [TimeSpan]::FromMinutes(5)
    $httpClient2.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $token)
    $httpClient2.DefaultRequestHeaders.Accept.Add([System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new("application/json"))

    try {
        $deployUrl = "$SiteUrl/_api/web/sitecollectionappcatalog/AvailableApps/GetById('$appId')/Deploy"
        $deployBody = [System.Net.Http.StringContent]::new('{"skipFeatureDeployment": true}', [System.Text.Encoding]::UTF8, "application/json")
        $deployResponse = $httpClient2.PostAsync($deployUrl, $deployBody).GetAwaiter().GetResult()

        if (-not $deployResponse.IsSuccessStatusCode) {
            $body = $deployResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            Write-Warning "Deploy may need manual approval: $($deployResponse.StatusCode)"
        } else {
            Write-Host "  Published" -ForegroundColor Green
        }
    } finally { $httpClient2.Dispose() }

    # Install app on site
    Write-Host "  Installing on site..." -ForegroundColor Yellow
    try {
        Install-PnPApp -Identity $appId -Scope Site -ErrorAction Stop
        Write-Host "  Installed" -ForegroundColor Green
    } catch {
        if ($_.Exception.Message -match "already installed|already exists") {
            Write-Host "  Already installed — updating..." -ForegroundColor Yellow
            Update-PnPApp -Identity $appId -Scope Site -ErrorAction SilentlyContinue
            Write-Host "  Updated" -ForegroundColor Green
        } else { throw }
    }

    Start-Sleep -Seconds 5  # Allow app to register

    # ═══════════════════════════════════════════════════════════════════
    # PHASE 3: Provision hidden lists
    # ═══════════════════════════════════════════════════════════════════
    Write-Host ""
    Write-Host "[3/5] Provisioning hidden lists..." -ForegroundColor Cyan

    $hiddenLists = @(
        @{
            Name = "SearchSavedQueries"
            Description = "SP Search saved queries and shared searches"
            Fields = @(
                @{ Name = "QueryText"; Type = "Note" },
                @{ Name = "QueryHash"; Type = "Text" },
                @{ Name = "SearchState"; Type = "Note" },
                @{ Name = "Vertical"; Type = "Text" },
                @{ Name = "Scope"; Type = "Text" },
                @{ Name = "IsShared"; Type = "Boolean" },
                @{ Name = "SharedWith"; Type = "Note" },
                @{ Name = "Tags"; Type = "Note" },
                @{ Name = "UseCount"; Type = "Number" },
                @{ Name = "LastUsed"; Type = "DateTime" }
            )
            Indexes = @("QueryHash", "Vertical")
        },
        @{
            Name = "SearchHistory"
            Description = "SP Search query history log"
            Fields = @(
                @{ Name = "QueryText"; Type = "Note" },
                @{ Name = "QueryHash"; Type = "Text" },
                @{ Name = "SearchState"; Type = "Note" },
                @{ Name = "Vertical"; Type = "Text" },
                @{ Name = "Scope"; Type = "Text" },
                @{ Name = "ResultCount"; Type = "Number" },
                @{ Name = "SearchTimestamp"; Type = "DateTime" },
                @{ Name = "ClickedUrls"; Type = "Note" }
            )
            Indexes = @("QueryHash", "SearchTimestamp", "Vertical")
        },
        @{
            Name = "SearchCollections"
            Description = "SP Search pinboards and collections"
            Fields = @(
                @{ Name = "CollectionName"; Type = "Text" },
                @{ Name = "CollectionDescription"; Type = "Note" },
                @{ Name = "ItemUrl"; Type = "Note" },
                @{ Name = "ItemTitle"; Type = "Text" },
                @{ Name = "ItemMetadata"; Type = "Note" },
                @{ Name = "IsShared"; Type = "Boolean" },
                @{ Name = "SharedWith"; Type = "Note" }
            )
            Indexes = @("CollectionName")
        }
    )

    foreach ($listDef in $hiddenLists) {
        $list = Get-PnPList -Identity $listDef.Name -ErrorAction SilentlyContinue
        if ($list) {
            Write-Host "  [EXISTS] $($listDef.Name)" -ForegroundColor Yellow
        } else {
            Write-Host "  [CREATE] $($listDef.Name)" -ForegroundColor Green
            $list = New-PnPList -Title $listDef.Name -Template GenericList -Hidden -EnableVersioning -ErrorAction Stop

            foreach ($field in $listDef.Fields) {
                Add-PnPField -List $listDef.Name -DisplayName $field.Name -InternalName $field.Name -Type $field.Type -ErrorAction SilentlyContinue | Out-Null
            }

            # Add indexes for query performance
            foreach ($idx in $listDef.Indexes) {
                try {
                    $f = Get-PnPField -List $listDef.Name -Identity $idx -ErrorAction SilentlyContinue
                    if ($f) {
                        Set-PnPField -List $listDef.Name -Identity $idx -Values @{ Indexed = $true } -ErrorAction SilentlyContinue
                    }
                } catch {
                    Write-Warning "  Could not index field '$idx' on $($listDef.Name)"
                }
            }

            # Set list as hidden
            Set-PnPList -Identity $listDef.Name -Hidden $true -ErrorAction SilentlyContinue
        }
    }
    Write-Host "  Hidden lists ready" -ForegroundColor Green

    # ═══════════════════════════════════════════════════════════════════
    # PHASE 4: Create search page with web parts
    # ═══════════════════════════════════════════════════════════════════
    Write-Host ""
    Write-Host "[4/5] Creating search page..." -ForegroundColor Cyan

    # Remove existing page if present
    $existingPage = Get-PnPPage -Identity $PageName -ErrorAction SilentlyContinue
    if ($existingPage) {
        Write-Host "  Removing existing page..." -ForegroundColor Yellow
        Remove-PnPPage -Identity $PageName -Force -ErrorAction Stop
    }

    Add-PnPPage -Name $PageName -Title "Search" -LayoutType Article -HeaderLayoutType NoImage -CommentsEnabled:$false -ErrorAction Stop | Out-Null
    Write-Host "  Page created" -ForegroundColor Green

    # Add sections
    Add-PnPPageSection -Page $PageName -SectionTemplate OneColumn -Order 1 -ErrorAction Stop         # Search Box
    Add-PnPPageSection -Page $PageName -SectionTemplate OneColumn -Order 2 -ErrorAction Stop         # Verticals
    Add-PnPPageSection -Page $PageName -SectionTemplate TwoColumnLeft -Order 3 -ErrorAction Stop     # Results | Filters
    Add-PnPPageSection -Page $PageName -SectionTemplate OneColumn -Order 4 -ErrorAction Stop         # Search Manager

    # Verticals config
    $defaultVerticals = @(
        @{ key = "all"; label = "All"; iconName = "Search"; sortOrder = 1 },
        @{ key = "documents"; label = "Documents"; iconName = "Page"; queryTemplate = "{searchTerms} contentclass:STS_ListItem_DocumentLibrary"; sortOrder = 2 },
        @{ key = "pages"; label = "Pages"; iconName = "FileHTML"; queryTemplate = "{searchTerms} (contentclass:STS_ListItem_WebPageLibrary OR contentclass:STS_Site)"; sortOrder = 3 },
        @{ key = "people"; label = "People"; iconName = "People"; queryTemplate = "{searchTerms}"; resultSourceId = "b09a7990-05ea-4af9-81ef-edfab16c4e31"; sortOrder = 4 },
        @{ key = "sites"; label = "Sites"; iconName = "Globe"; queryTemplate = "{searchTerms} contentclass:STS_Site"; sortOrder = 5 }
    ) | ConvertTo-Json -Depth 4 -Compress

    # Sortable properties config
    $defaultSortableProperties = @(
        @{ property = "LastModifiedTime"; label = "Date (newest)"; direction = "Descending" },
        @{ property = "LastModifiedTime"; label = "Date (oldest)"; direction = "Ascending" },
        @{ property = "Size"; label = "Size (largest)"; direction = "Descending" },
        @{ property = "ViewsLifeTime"; label = "Popularity"; direction = "Descending" }
    ) | ConvertTo-Json -Depth 4 -Compress

    # Filters config - configured for the test data site columns
    $defaultFilters = @(
        @{ uniqueId = "f1"; managedProperty = "FileType"; displayName = "File Type"; filterType = "checkbox"; operator = "OR"; maxValues = 15; showCount = $true; defaultExpanded = $true; sortBy = "count" },
        @{ uniqueId = "f2"; managedProperty = "Author"; displayName = "Author"; filterType = "people"; operator = "OR"; maxValues = 10; showCount = $true; defaultExpanded = $true; sortBy = "count" },
        @{ uniqueId = "f3"; managedProperty = "LastModifiedTime"; displayName = "Modified Date"; filterType = "daterange"; operator = "AND"; maxValues = 1; showCount = $false; defaultExpanded = $true; sortBy = "alphabetical" },
        @{ uniqueId = "f4"; managedProperty = "contentclass"; displayName = "Content Type"; filterType = "checkbox"; operator = "OR"; maxValues = 10; showCount = $true; defaultExpanded = $false; sortBy = "count" }
    ) | ConvertTo-Json -Depth 4 -Compress

    # Add web parts
    Write-Host "  Adding Search Box..." -ForegroundColor Yellow
    Add-PnPPageWebPart -Page $PageName -Component $WP_SEARCH_BOX -Section 1 -Column 1 -WebPartProperties @{
        searchContextId     = $SearchContextId
        placeholder         = "Search SharePoint..."
        debounceMs          = 300
        searchBehavior      = "both"
        enableScopeSelector = $true
        enableSuggestions   = $true
        enableSearchManager = $false
        enableQueryBuilder  = $false
    } -ErrorAction Stop | Out-Null
    Write-Host "  [OK] Search Box" -ForegroundColor Green

    Write-Host "  Adding Verticals..." -ForegroundColor Yellow
    Add-PnPPageWebPart -Page $PageName -Component $WP_SEARCH_VERTICALS -Section 2 -Column 1 -WebPartProperties @{
        searchContextId    = $SearchContextId
        verticals          = $defaultVerticals
        showCounts         = $true
        hideEmptyVerticals = $false
        tabStyle           = "underline"
    } -ErrorAction Stop | Out-Null
    Write-Host "  [OK] Verticals (All, Documents, Pages, People, Sites)" -ForegroundColor Green

    Write-Host "  Adding Search Results..." -ForegroundColor Yellow
    Add-PnPPageWebPart -Page $PageName -Component $WP_SEARCH_RESULTS -Section 3 -Column 1 -WebPartProperties @{
        searchContextId              = $SearchContextId
        queryTemplate                = "{searchTerms}"
        pageSize                     = 25
        defaultLayout                = "list"
        showResultCount              = $true
        showSortDropdown             = $true
        enableSelection              = $true
        sortablePropertiesCollection = $defaultSortableProperties
    } -ErrorAction Stop | Out-Null
    Write-Host "  [OK] Search Results (List layout, 25/page, 4 sort options)" -ForegroundColor Green

    Write-Host "  Adding Search Filters..." -ForegroundColor Yellow
    Add-PnPPageWebPart -Page $PageName -Component $WP_SEARCH_FILTERS -Section 3 -Column 2 -WebPartProperties @{
        searchContextId        = $SearchContextId
        filtersCollection      = $defaultFilters
        applyMode              = "instant"
        operatorBetweenFilters = "AND"
        showClearAll           = $true
        enableVisualFilterBuilder = $false
    } -ErrorAction Stop | Out-Null
    Write-Host "  [OK] Search Filters (FileType, Author, Date, ContentClass)" -ForegroundColor Green

    Write-Host "  Adding Search Manager..." -ForegroundColor Yellow
    Add-PnPPageWebPart -Page $PageName -Component $WP_SEARCH_MANAGER -Section 4 -Column 1 -WebPartProperties @{
        searchContextId = $SearchContextId
        mode            = "standalone"
    } -ErrorAction Stop | Out-Null
    Write-Host "  [OK] Search Manager (standalone)" -ForegroundColor Green

    # ═══════════════════════════════════════════════════════════════════
    # PHASE 5: Publish
    # ═══════════════════════════════════════════════════════════════════
    Write-Host ""
    Write-Host "[5/5] Publishing page..." -ForegroundColor Cyan
    Set-PnPPage -Identity $PageName -Publish -ErrorAction Stop
    Write-Host "  Published" -ForegroundColor Green

    # ═══════════════════════════════════════════════════════════════════
    # Summary
    # ═══════════════════════════════════════════════════════════════════
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host " Setup Complete!" -ForegroundColor Green
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host ""
    Write-Host "  App deployed:     sp-search.sppkg ($packageSize MB)" -ForegroundColor White
    Write-Host "  Hidden lists:     SearchSavedQueries, SearchHistory, SearchCollections" -ForegroundColor White
    Write-Host "  Search page:      $SiteUrl/SitePages/$PageName.aspx" -ForegroundColor White
    Write-Host ""
    Write-Host "  Page Layout:" -ForegroundColor Yellow
    Write-Host "  ┌──────────────────────────────────────────────────────┐"
    Write-Host "  │ Search Box (full width)                              │"
    Write-Host "  ├──────────────────────────────────────────────────────┤"
    Write-Host "  │ All │ Documents │ Pages │ People │ Sites             │"
    Write-Host "  ├────────────────────────────┬─────────────────────────┤"
    Write-Host "  │ Results (66%)              │ Filters (33%)          │"
    Write-Host "  │  - List layout, 25/page    │  - File Type           │"
    Write-Host "  │  - 4 sort options          │  - Author (people)     │"
    Write-Host "  │  - Selection enabled       │  - Modified Date       │"
    Write-Host "  │                            │  - Content Type        │"
    Write-Host "  ├────────────────────────────┴─────────────────────────┤"
    Write-Host "  │ Search Manager (saved searches, history, collections)│"
    Write-Host "  └──────────────────────────────────────────────────────┘"
    Write-Host ""
    Write-Host "  Web Part Properties Set:" -ForegroundColor Yellow
    Write-Host "    Search Box:     contextId=$SearchContextId, debounce=300ms, suggestions=on"
    Write-Host "    Verticals:      5 tabs (All, Documents, Pages, People, Sites)"
    Write-Host "    Results:        list layout, 25/page, sort dropdown, selection"
    Write-Host "    Filters:        4 refiners (FileType, Author, Date, ContentClass)"
    Write-Host "    Sort options:   Date newest/oldest, Size, Popularity"
    Write-Host "    Search Manager: standalone mode"
    Write-Host ""
    Write-Host "  Open: $SiteUrl/SitePages/$PageName.aspx" -ForegroundColor Cyan
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host " Setup failed!" -ForegroundColor Red
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Stack: $($_.ScriptStackTrace)" -ForegroundColor Yellow
    Write-Host ""
    exit 1

} finally {
    try {
        $null = Get-PnPConnection -ErrorAction Stop
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
    } catch { }
}
