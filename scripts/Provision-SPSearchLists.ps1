<#
.SYNOPSIS
    Provisions the three hidden SharePoint lists required by SP Search.

.DESCRIPTION
    Creates SearchSavedQueries, SearchHistory, and SearchCollections
    hidden lists with proper columns, indexes, and permissions.
    Idempotent — safe to re-run.

.PARAMETER SiteUrl
    The SharePoint site collection URL where lists will be created.

.PARAMETER ClientId
    Azure AD app registration Client ID for PnP authentication.
    If omitted, falls back to default PnP interactive login.


.EXAMPLE
    .\Provision-SPSearchLists.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/search" -ClientId "970bb320-0d49-4b4a-aa8f-c3f4b1e5928f"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$ClientId
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
if ($ClientId) {
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
} else {
    Connect-PnPOnline -Url $SiteUrl -Interactive
}

# ─── 1. SearchSavedQueries ────────────────────────────────
Write-Host "`n[1/3] SearchSavedQueries" -ForegroundColor Magenta

Ensure-HiddenList -ListName "SearchSavedQueries" -Description "SP Search: Saved/shared searches and state snapshots"

# Columns
Ensure-Field -ListName "SearchSavedQueries" -FieldName "QueryText" -FieldType "Note"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "SearchState" -FieldType "Note"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "SearchUrl" -FieldType "URL"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "EntryType" -FieldType "Choice" -Choices @("SavedSearch", "SharedSearch", "StateSnapshot")
Ensure-Field -ListName "SearchSavedQueries" -FieldName "Category" -FieldType "Text"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "SharedWith" -FieldType "UserMulti"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "ResultCount" -FieldType "Number"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "LastUsed" -FieldType "DateTime"
Ensure-Field -ListName "SearchSavedQueries" -FieldName "ExpiresAt" -FieldType "DateTime"

# Indexes
Ensure-Index -ListName "SearchSavedQueries" -FieldName "Title"
Ensure-Index -ListName "SearchSavedQueries" -FieldName "EntryType"
Ensure-Index -ListName "SearchSavedQueries" -FieldName "Category"
Ensure-Index -ListName "SearchSavedQueries" -FieldName "LastUsed"
Ensure-Index -ListName "SearchSavedQueries" -FieldName "ExpiresAt"

# Permissions: All authenticated users can Add Items
Write-Host "  [PERM] Setting permissions for SearchSavedQueries..." -ForegroundColor Cyan
Set-PnPList -Identity "SearchSavedQueries" -BreakRoleInheritance -CopyRoleAssignments

# ─── 2. SearchHistory ─────────────────────────────────────
Write-Host "`n[2/3] SearchHistory" -ForegroundColor Magenta

Ensure-HiddenList -ListName "SearchHistory" -Description "SP Search: User search history (high-volume, auto-pruned)"

# Columns
Ensure-Field -ListName "SearchHistory" -FieldName "QueryText" -FieldType "Note"
Ensure-Field -ListName "SearchHistory" -FieldName "QueryHash" -FieldType "Text"
Ensure-Field -ListName "SearchHistory" -FieldName "Vertical" -FieldType "Text"
Ensure-Field -ListName "SearchHistory" -FieldName "Scope" -FieldType "Text"
Ensure-Field -ListName "SearchHistory" -FieldName "SearchState" -FieldType "Note"
Ensure-Field -ListName "SearchHistory" -FieldName "ResultCount" -FieldType "Number"
Ensure-Field -ListName "SearchHistory" -FieldName "IsZeroResult" -FieldType "Boolean"
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
Write-Host "`n[3/3] SearchCollections" -ForegroundColor Magenta

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

# ─── Summary ──────────────────────────────────────────────
Write-Host "`n================================================" -ForegroundColor Cyan
Write-Host " Provisioning Complete!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "`nCreated/verified:"
Write-Host "  - SearchSavedQueries (saved + shared searches)"
Write-Host "  - SearchHistory (user history, indexed for >5K items)"
Write-Host "  - SearchCollections (result pinboards)"

Disconnect-PnPOnline
