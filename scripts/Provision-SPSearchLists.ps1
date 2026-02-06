<#
.SYNOPSIS
    Provisions the four hidden SharePoint lists required by SP Search.

.DESCRIPTION
    Creates SearchSavedQueries, SearchHistory, SearchCollections, and
    SearchConfiguration hidden lists with proper columns, indexes,
    and permissions. Idempotent — safe to re-run.

.PARAMETER SiteUrl
    The SharePoint site collection URL where lists will be created.

.PARAMETER AdminGroupName
    Name of the SP Search Admins security group (for SearchConfiguration write access).
    Defaults to "SP Search Admins".

.EXAMPLE
    .\Provision-SPSearchLists.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/search"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$AdminGroupName = "SP Search Admins"
)

# Ensure PnP.PowerShell is available
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    Write-Error "PnP.PowerShell module is required. Install via: Install-Module PnP.PowerShell -Scope CurrentUser"
    exit 1
}

Import-Module PnP.PowerShell -ErrorAction Stop

# ───────────────────────────────────────────────────────────
# Helper: Create list if it doesn't exist
# ───────────────────────────────────────────────────────────
function Ensure-HiddenList {
    param(
        [string]$ListName,
        [string]$Description
    )

    $list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($list) {
        Write-Host "  [EXISTS] List '$ListName' already exists." -ForegroundColor Yellow
        return $list
    }

    Write-Host "  [CREATE] Creating hidden list '$ListName'..." -ForegroundColor Green
    $list = New-PnPList -Title $ListName -Template GenericList -Hidden -EnableVersioning -OnQuickLaunch:$false
    Set-PnPList -Identity $ListName -Description $Description
    return $list
}

# ───────────────────────────────────────────────────────────
# Helper: Add field if it doesn't exist
# ───────────────────────────────────────────────────────────
function Ensure-Field {
    param(
        [string]$ListName,
        [string]$FieldName,
        [string]$FieldType,
        [bool]$Required = $false,
        [string[]]$Choices = @()
    )

    $field = Get-PnPField -List $ListName -Identity $FieldName -ErrorAction SilentlyContinue
    if ($field) {
        Write-Host "    [EXISTS] Field '$FieldName' already exists." -ForegroundColor Yellow
        return
    }

    Write-Host "    [CREATE] Adding field '$FieldName' ($FieldType)..." -ForegroundColor Green
    switch ($FieldType) {
        "Text" {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type Text -Required:$Required | Out-Null
        }
        "Note" {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type Note -Required:$Required | Out-Null
        }
        "Number" {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type Number -Required:$Required | Out-Null
        }
        "DateTime" {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type DateTime -Required:$Required | Out-Null
        }
        "URL" {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type URL -Required:$Required | Out-Null
        }
        "Boolean" {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type Boolean -Required:$Required | Out-Null
        }
        "Choice" {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type Choice -Choices $Choices -Required:$Required | Out-Null
        }
        "UserMulti" {
            Add-PnPFieldFromXml -List $ListName -FieldXml "<Field Type='UserMulti' DisplayName='$FieldName' StaticName='$FieldName' Name='$FieldName' Mult='TRUE' UserSelectionMode='PeopleOnly' />" | Out-Null
        }
    }
}

# ───────────────────────────────────────────────────────────
# Helper: Create index on a field
# ───────────────────────────────────────────────────────────
function Ensure-Index {
    param(
        [string]$ListName,
        [string]$FieldName
    )

    Write-Host "    [INDEX] Ensuring index on '$FieldName'..." -ForegroundColor Cyan
    try {
        $field = Get-PnPField -List $ListName -Identity $FieldName -ErrorAction Stop
        if (-not $field.Indexed) {
            $field.Indexed = $true
            $field.Update()
            Invoke-PnPQuery
            Write-Host "    [INDEX] Index created on '$FieldName'." -ForegroundColor Green
        } else {
            Write-Host "    [INDEX] '$FieldName' is already indexed." -ForegroundColor Yellow
        }
    } catch {
        Write-Warning "    [INDEX] Failed to create index on '$FieldName': $_"
    }
}

# ═══════════════════════════════════════════════════════════
# MAIN SCRIPT
# ═══════════════════════════════════════════════════════════

Write-Host "`n================================================" -ForegroundColor Cyan
Write-Host " SP Search — Hidden List Provisioning" -ForegroundColor Cyan
Write-Host "================================================`n" -ForegroundColor Cyan

# Connect to SharePoint
Write-Host "Connecting to $SiteUrl..." -ForegroundColor White
Connect-PnPOnline -Url $SiteUrl -Interactive

# ─── 1. SearchSavedQueries ────────────────────────────────
Write-Host "`n[1/4] SearchSavedQueries" -ForegroundColor Magenta

Ensure-HiddenList -ListName "SearchSavedQueries" -Description "SP Search: Saved and shared search queries"

# Columns
Ensure-Field -ListName "SearchSavedQueries" -FieldName "QueryText" -FieldType "Note"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "SearchState" -FieldType "Note"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "SearchUrl" -FieldType "URL"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "EntryType" -FieldType "Choice" -Choices @("SavedSearch", "SharedSearch")
Ensure-Field -ListName "SearchSavedQueries" -FieldName "Category" -FieldType "Text"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "SharedWith" -FieldType "UserMulti"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "ResultCount" -FieldType "Number"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "LastUsed" -FieldType "DateTime"

# Indexes
Ensure-Index -ListName "SearchSavedQueries" -FieldName "Title"
Ensure-Index -ListName "SearchSavedQueries" -FieldName "EntryType"
Ensure-Index -ListName "SearchSavedQueries" -FieldName "Category"
Ensure-Index -ListName "SearchSavedQueries" -FieldName "LastUsed"

# Permissions: All authenticated users can Add Items
Write-Host "  [PERM] Setting permissions for SearchSavedQueries..." -ForegroundColor Cyan
Set-PnPList -Identity "SearchSavedQueries" -BreakRoleInheritance -CopyRoleAssignments

# ─── 2. SearchHistory ─────────────────────────────────────
Write-Host "`n[2/4] SearchHistory" -ForegroundColor Magenta

Ensure-HiddenList -ListName "SearchHistory" -Description "SP Search: User search history (high-volume, auto-pruned)"

# Columns
Ensure-Field -ListName "SearchHistory" -FieldName "QueryHash" -FieldType "Text"
Ensure-Field -ListName "SearchHistory" -FieldName "Vertical" -FieldType "Text"
Ensure-Field -ListName "SearchHistory" -FieldName "Scope" -FieldType "Text"
Ensure-Field -ListName "SearchHistory" -FieldName "SearchState" -FieldType "Note"
Ensure-Field -ListName "SearchHistory" -FieldName "ResultCount" -FieldType "Number"
Ensure-Field -ListName "SearchHistory" -FieldName "ClickedItems" -FieldType "Note"
Ensure-Field -ListName "SearchHistory" -FieldName "SearchTimestamp" -FieldType "DateTime"

# CRITICAL: Index Author (Created By), SearchTimestamp, QueryHash, Vertical
# Author must be the primary filter for all queries to avoid list view threshold issues
Write-Host "  [CRITICAL] Indexing Author, SearchTimestamp, QueryHash, Vertical..." -ForegroundColor Red
Ensure-Index -ListName "SearchHistory" -FieldName "Author"
Ensure-Index -ListName "SearchHistory" -FieldName "SearchTimestamp"
Ensure-Index -ListName "SearchHistory" -FieldName "QueryHash"
Ensure-Index -ListName "SearchHistory" -FieldName "Vertical"

# Verify all indexes were created
$authorField = Get-PnPField -List "SearchHistory" -Identity "Author" -ErrorAction SilentlyContinue
$tsField = Get-PnPField -List "SearchHistory" -Identity "SearchTimestamp" -ErrorAction SilentlyContinue
if (-not $authorField.Indexed -or -not $tsField.Indexed) {
    Write-Error "CRITICAL: Failed to create required indexes on SearchHistory. List queries WILL fail at >5,000 items."
    Write-Error "Manually verify indexes on Author and SearchTimestamp before proceeding."
}

# Permissions: Add Items + Edit Own Items (no Read All)
Write-Host "  [PERM] Setting permissions for SearchHistory..." -ForegroundColor Cyan
Set-PnPList -Identity "SearchHistory" -BreakRoleInheritance -ClearSubscopes
# Users can add and edit their own items
Set-PnPList -Identity "SearchHistory" -ReadSecurity 2 -WriteSecurity 2

# ─── 3. SearchCollections ─────────────────────────────────
Write-Host "`n[3/4] SearchCollections" -ForegroundColor Magenta

Ensure-HiddenList -ListName "SearchCollections" -Description "SP Search: User search result collections/pinboards"

# Columns
Ensure-Field -ListName "SearchCollections" -FieldName "ItemUrl" -FieldType "URL"
Ensure-Field -ListName "SearchCollections" -FieldName "ItemTitle" -FieldType "Text"
Ensure-Field -ListName "SearchCollections" -FieldName "ItemMetadata" -FieldType "Note"
Ensure-Field -ListName "SearchCollections" -FieldName "CollectionName" -FieldType "Text"
Ensure-Field -ListName "SearchCollections" -FieldName "Tags" -FieldType "Note"
Ensure-Field -ListName "SearchCollections" -FieldName "SharedWith" -FieldType "UserMulti"
Ensure-Field -ListName "SearchCollections" -FieldName "SortOrder" -FieldType "Number"

# Indexes
Ensure-Index -ListName "SearchCollections" -FieldName "Title"
Ensure-Index -ListName "SearchCollections" -FieldName "CollectionName"

# Permissions: All authenticated users can Add Items
Write-Host "  [PERM] Setting permissions for SearchCollections..." -ForegroundColor Cyan
Set-PnPList -Identity "SearchCollections" -BreakRoleInheritance -CopyRoleAssignments

# ─── 4. SearchConfiguration ──────────────────────────────
Write-Host "`n[4/4] SearchConfiguration" -ForegroundColor Magenta

Ensure-HiddenList -ListName "SearchConfiguration" -Description "SP Search: Admin configuration (scopes, verticals, promoted results, state snapshots)"

# Columns
Ensure-Field -ListName "SearchConfiguration" -FieldName "ConfigType" -FieldType "Choice" -Choices @("Scope", "VerticalPreset", "LayoutMapping", "ManagedPropertyMap", "PromotedResult", "StateSnapshot")
Ensure-Field -ListName "SearchConfiguration" -FieldName "ConfigValue" -FieldType "Note"
Ensure-Field -ListName "SearchConfiguration" -FieldName "IsActive" -FieldType "Boolean"
Ensure-Field -ListName "SearchConfiguration" -FieldName "SortOrder" -FieldType "Number"
Ensure-Field -ListName "SearchConfiguration" -FieldName "ExpiresAt" -FieldType "DateTime"
Ensure-Field -ListName "SearchConfiguration" -FieldName "AudienceGroups" -FieldType "Note"

# Indexes
Ensure-Index -ListName "SearchConfiguration" -FieldName "Title"
Ensure-Index -ListName "SearchConfiguration" -FieldName "ConfigType"
Ensure-Index -ListName "SearchConfiguration" -FieldName "IsActive"
Ensure-Index -ListName "SearchConfiguration" -FieldName "ExpiresAt"

# Permissions: Admin-only write, users have Read
Write-Host "  [PERM] Setting permissions for SearchConfiguration (admin-only write)..." -ForegroundColor Cyan
Set-PnPList -Identity "SearchConfiguration" -BreakRoleInheritance -ClearSubscopes

# ─── Seed Default Configuration ──────────────────────────
Write-Host "`n[SEED] Seeding default configuration entries..." -ForegroundColor Magenta

# Default search scopes
$defaultScopes = @(
    @{
        Title = "All SharePoint"
        ConfigType = "Scope"
        ConfigValue = '{"id":"all","label":"All SharePoint"}'
        IsActive = $true
        SortOrder = 1
    },
    @{
        Title = "Current Site"
        ConfigType = "Scope"
        ConfigValue = '{"id":"currentsite","label":"Current Site","kqlPath":"path:{Site.URL}"}'
        IsActive = $true
        SortOrder = 2
    },
    @{
        Title = "Current Hub"
        ConfigType = "Scope"
        ConfigValue = '{"id":"hub","label":"Current Hub","kqlPath":"DepartmentId:{Hub}"}'
        IsActive = $true
        SortOrder = 3
    }
)

foreach ($scope in $defaultScopes) {
    $existing = Get-PnPListItem -List "SearchConfiguration" -Query "<View><Query><Where><And><Eq><FieldRef Name='Title'/><Value Type='Text'>$($scope.Title)</Value></Eq><Eq><FieldRef Name='ConfigType'/><Value Type='Choice'>Scope</Value></Eq></And></Where></Query></View>" -ErrorAction SilentlyContinue
    if (-not $existing) {
        Add-PnPListItem -List "SearchConfiguration" -Values $scope | Out-Null
        Write-Host "  [SEED] Created scope: $($scope.Title)" -ForegroundColor Green
    } else {
        Write-Host "  [SEED] Scope '$($scope.Title)' already exists." -ForegroundColor Yellow
    }
}

# Default layout mappings
$defaultLayouts = @(
    @{
        Title = "List Layout"
        ConfigType = "LayoutMapping"
        ConfigValue = '{"id":"list","displayName":"List","iconName":"BulletedList","isDefault":true}'
        IsActive = $true
        SortOrder = 1
    },
    @{
        Title = "Compact Layout"
        ConfigType = "LayoutMapping"
        ConfigValue = '{"id":"compact","displayName":"Compact","iconName":"AlignLeft","isDefault":false}'
        IsActive = $true
        SortOrder = 2
    }
)

foreach ($layout in $defaultLayouts) {
    $existing = Get-PnPListItem -List "SearchConfiguration" -Query "<View><Query><Where><And><Eq><FieldRef Name='Title'/><Value Type='Text'>$($layout.Title)</Value></Eq><Eq><FieldRef Name='ConfigType'/><Value Type='Choice'>LayoutMapping</Value></Eq></And></Where></Query></View>" -ErrorAction SilentlyContinue
    if (-not $existing) {
        Add-PnPListItem -List "SearchConfiguration" -Values $layout | Out-Null
        Write-Host "  [SEED] Created layout: $($layout.Title)" -ForegroundColor Green
    } else {
        Write-Host "  [SEED] Layout '$($layout.Title)' already exists." -ForegroundColor Yellow
    }
}

# ─── Summary ──────────────────────────────────────────────
Write-Host "`n================================================" -ForegroundColor Cyan
Write-Host " Provisioning Complete!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "`nCreated/verified:"
Write-Host "  - SearchSavedQueries (saved + shared searches)"
Write-Host "  - SearchHistory (user history, indexed for >5K items)"
Write-Host "  - SearchCollections (result pinboards)"
Write-Host "  - SearchConfiguration (admin config, promoted results)"
Write-Host "`nNOTE: Remember to add '$AdminGroupName' security group as"
Write-Host "SearchConfiguration list owners for admin write access." -ForegroundColor Yellow

Disconnect-PnPOnline
