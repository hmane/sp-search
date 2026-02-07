<#
.SYNOPSIS
    Maps crawled properties to managed properties for SP Search test data.

.DESCRIPTION
    Uses the SharePoint search configuration XML import API to map custom
    crawled properties (ows_SPS*) to pre-provisioned managed properties
    (RefinableStringXX, RefinableDateXX, etc.) at the site collection level.

    This script should be run AFTER the first search crawl completes so the
    crawled properties exist. After mapping, request a re-index for the
    mapped properties to take effect.

    Based on Mikael Svenson's SearchConfiguration template approach.

.PARAMETER SiteUrl
    The SharePoint site collection URL (e.g., https://pixelboy.sharepoint.com/sites/SPSearch)

.PARAMETER ClientId
    Azure AD application client ID for authentication.

.PARAMETER RequestReindex
    Request a site re-index after applying mappings.

.PARAMETER SkipValidation
    Skip crawled property existence validation.

.EXAMPLE
    .\Map-CrawledProperties.ps1 -SiteUrl "https://pixelboy.sharepoint.com/sites/SPSearch"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [string]$ClientId,

    [switch]$RequestReindex,

    [switch]$SkipValidation
)

$ErrorActionPreference = 'Stop'

# ═══════════════════════════════════════════════════════════════════════════
# Mapping Table
# ═══════════════════════════════════════════════════════════════════════════

# Propset GUIDs
$PROPSET_SHAREPOINT = '00130329-0000-0130-c000-000000131346'  # Standard ows_* columns
$PROPSET_TAXONOMY   = '158d7563-aeff-4dbf-bf16-4a1445f0366c'  # ows_taxId_* taxonomy columns

$mappings = @(
    @{ CrawledProperty = 'ows_SPSStatus';           ManagedProperty = 'RefinableString00';  Propset = $PROPSET_SHAREPOINT; FilterType = 'Checkbox' }
    @{ CrawledProperty = 'ows_SPSPriority';          ManagedProperty = 'RefinableString01';  Propset = $PROPSET_SHAREPOINT; FilterType = 'Checkbox' }
    @{ CrawledProperty = 'ows_SPSRegion';            ManagedProperty = 'RefinableString02';  Propset = $PROPSET_SHAREPOINT; FilterType = 'TagBox' }
    @{ CrawledProperty = 'ows_SPSDocumentType';      ManagedProperty = 'RefinableString03';  Propset = $PROPSET_SHAREPOINT; FilterType = 'TagBox' }
    @{ CrawledProperty = 'ows_taxId_SPSDepartment';  ManagedProperty = 'RefinableString05';  Propset = $PROPSET_TAXONOMY;   FilterType = 'Taxonomy' }
    @{ CrawledProperty = 'ows_taxId_SPSTags';        ManagedProperty = 'RefinableString06';  Propset = $PROPSET_TAXONOMY;   FilterType = 'TagBox (multi)' }
    @{ CrawledProperty = 'ows_SPSIsActive';          ManagedProperty = 'RefinableString07';  Propset = $PROPSET_SHAREPOINT; FilterType = 'Toggle' }
    @{ CrawledProperty = 'ows_SPSIsPublished';       ManagedProperty = 'RefinableString08';  Propset = $PROPSET_SHAREPOINT; FilterType = 'Toggle' }
    @{ CrawledProperty = 'ows_SPSOwner';             ManagedProperty = 'RefinableString09';  Propset = $PROPSET_SHAREPOINT; FilterType = 'People' }
    @{ CrawledProperty = 'ows_SPSBudget';            ManagedProperty = 'RefinableDecimal00'; Propset = $PROPSET_SHAREPOINT; FilterType = 'Slider' }
    @{ CrawledProperty = 'ows_SPSRating';            ManagedProperty = 'RefinableDecimal01'; Propset = $PROPSET_SHAREPOINT; FilterType = 'Slider' }
    @{ CrawledProperty = 'ows_SPSViewCount';         ManagedProperty = 'RefinableInt00';     Propset = $PROPSET_SHAREPOINT; FilterType = 'Number/Sort' }
    @{ CrawledProperty = 'ows_SPSReviewDate';        ManagedProperty = 'RefinableDate00';    Propset = $PROPSET_SHAREPOINT; FilterType = 'DateRange' }
)

# ═══════════════════════════════════════════════════════════════════════════
# PID Calculation
# ═══════════════════════════════════════════════════════════════════════════

function Get-ManagedPid {
    param([string]$ManagedProperty)

    $match = [regex]::Match($ManagedProperty, '^(?<prefix>Refinable(?:String|Date|DateSingle|DateInvariant|Decimal|Double|Int)|Double|Decimal|Date|Int)(?<num>\d+)$')
    if (-not $match.Success) {
        throw "Unsupported managed property name: $ManagedProperty"
    }

    $prefix = $match.Groups['prefix'].Value
    $num = [int]$match.Groups['num'].Value

    switch -Regex ($prefix) {
        '^RefinableString$' {
            if ($num -ge 100) { return 1000000900 + ($num - 100) }
            return 1000000000 + $num
        }
        '^RefinableDouble$'        { return 1000000800 + $num }
        '^RefinableDecimal$'       { return 1000000700 + $num }
        '^RefinableDateInvariant$' { return 1000000660 + $num }
        '^RefinableDateSingle$'    { return 1000000660 + $num }
        '^RefinableDate$'          { return 1000000600 + $num }
        '^RefinableInt$'           { return 1000000500 + $num }
        '^Double$'                 { return 1000000400 + $num }
        '^Decimal$'                { return 1000000300 + $num }
        '^Date$'                   { return 1000000200 + $num }
        '^Int$'                    { return 1000000100 + $num }
        default { throw "Cannot determine base PID for: $ManagedProperty" }
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# Search Configuration XML Template (Mikael Svenson / wobba approach)
# ═══════════════════════════════════════════════════════════════════════════

# Minimal search config XML — imports only the mapping, not the full schema.
# ManagedProperties section has TotalCount=0 because the RefinableXX properties
# already exist at the tenant level; we are only adding mappings to them.
function Get-MappingXml {
    param(
        [string]$CrawledPropertyName,
        [string]$PropsetGuid,
        [int]$ManagedPid
    )

    return @"
<SearchConfigurationSettings xmlns:i="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.datacontract.org/2004/07/Microsoft.Office.Server.Search.Portability">
  <SearchQueryConfigurationSettings>
    <SearchQueryConfigurationSettings>
      <BestBets xmlns:d4p1="http://www.microsoft.com/sharepoint/search/KnownTypes/2008/08" />
      <DefaultSourceId>00000000-0000-0000-0000-000000000000</DefaultSourceId>
      <DefaultSourceIdSet>true</DefaultSourceIdSet>
      <DeployToParent>false</DeployToParent>
      <DisableInheritanceOnImport>false</DisableInheritanceOnImport>
      <QueryRuleGroups xmlns:d4p1="http://www.microsoft.com/sharepoint/search/KnownTypes/2008/08" />
      <QueryRules xmlns:d4p1="http://www.microsoft.com/sharepoint/search/KnownTypes/2008/08" />
      <ResultTypes xmlns:d4p1="http://schemas.datacontract.org/2004/07/Microsoft.Office.Server.Search.Administration" />
      <Sources xmlns:d4p1="http://schemas.datacontract.org/2004/07/Microsoft.Office.Server.Search.Administration.Query" />
      <UserSegments xmlns:d4p1="http://www.microsoft.com/sharepoint/search/KnownTypes/2008/08" />
    </SearchQueryConfigurationSettings>
  </SearchQueryConfigurationSettings>
  <SearchRankingModelConfigurationSettings>
    <RankingModels xmlns:d3p1="http://schemas.microsoft.com/2003/10/Serialization/Arrays" />
  </SearchRankingModelConfigurationSettings>
  <SearchSchemaConfigurationSettings>
    <Aliases xmlns:d3p1="http://schemas.datacontract.org/2004/07/Microsoft.Office.Server.Search.Administration">
      <d3p1:LastItemName i:nil="true" />
      <d3p1:dictionary xmlns:d4p1="http://schemas.microsoft.com/2003/10/Serialization/Arrays" />
    </Aliases>
    <CategoriesAndCrawledProperties xmlns:d3p1="http://schemas.microsoft.com/2003/10/Serialization/Arrays">
      <d3p1:KeyValueOfguidCrawledPropertyInfoCollectionaSYUqUE_P>
        <d3p1:Key>${PropsetGuid}</d3p1:Key>
        <d3p1:Value xmlns:d5p1="http://schemas.datacontract.org/2004/07/Microsoft.Office.Server.Search.Administration">
          <d5p1:LastItemName>${CrawledPropertyName}</d5p1:LastItemName>
          <d5p1:dictionary>
            <d3p1:KeyValueOfstringCrawledPropertyInfoy6h3NzC8>
              <d3p1:Key>${CrawledPropertyName}</d3p1:Key>
              <d3p1:Value>
                <d5p1:CategoryName>SharePoint</d5p1:CategoryName>
                <d5p1:IsImplicit>false</d5p1:IsImplicit>
                <d5p1:IsMappedToContents>true</d5p1:IsMappedToContents>
                <d5p1:IsNameEnum>false</d5p1:IsNameEnum>
                <d5p1:MappedManagedProperties />
                <d5p1:Name>${CrawledPropertyName}</d5p1:Name>
                <d5p1:Propset>${PropsetGuid}</d5p1:Propset>
                <d5p1:Samples />
                <d5p1:SchemaId>143692</d5p1:SchemaId>
              </d3p1:Value>
            </d3p1:KeyValueOfstringCrawledPropertyInfoy6h3NzC8>
          </d5p1:dictionary>
        </d3p1:Value>
      </d3p1:KeyValueOfguidCrawledPropertyInfoCollectionaSYUqUE_P>
    </CategoriesAndCrawledProperties>
    <CrawledProperties xmlns:d3p1="http://schemas.datacontract.org/2004/07/Microsoft.Office.Server.Search.Administration">
      <d3p1:LastItemName i:nil="true" />
      <d3p1:dictionary xmlns:d4p1="http://schemas.microsoft.com/2003/10/Serialization/Arrays" />
    </CrawledProperties>
    <ManagedProperties xmlns:d3p1="http://schemas.datacontract.org/2004/07/Microsoft.Office.Server.Search.Administration">
      <d3p1:LastItemName i:nil="true" />
      <d3p1:dictionary xmlns:d4p1="http://schemas.microsoft.com/2003/10/Serialization/Arrays" />
      <d3p1:TotalCount>0</d3p1:TotalCount>
    </ManagedProperties>
    <Mappings xmlns:d3p1="http://schemas.datacontract.org/2004/07/Microsoft.Office.Server.Search.Administration">
      <d3p1:LastItemName i:nil="true" />
      <d3p1:dictionary xmlns:d4p1="http://schemas.microsoft.com/2003/10/Serialization/Arrays">
        <d4p1:KeyValueOfstringMappingInfoy6h3NzC8>
          <d4p1:Key>${PropsetGuid}:${CrawledPropertyName}-&gt;${ManagedPid}</d4p1:Key>
          <d4p1:Value>
            <d3p1:CrawledPropertyName>${CrawledPropertyName}</d3p1:CrawledPropertyName>
            <d3p1:CrawledPropset>${PropsetGuid}</d3p1:CrawledPropset>
            <d3p1:ManagedPid>${ManagedPid}</d3p1:ManagedPid>
            <d3p1:MappingOrder>100</d3p1:MappingOrder>
            <d3p1:Name>${PropsetGuid}:${CrawledPropertyName}-&gt;${ManagedPid}</d3p1:Name>
            <d3p1:SchemaId>143692</d3p1:SchemaId>
          </d4p1:Value>
        </d4p1:KeyValueOfstringMappingInfoy6h3NzC8>
      </d3p1:dictionary>
    </Mappings>
    <Overrides xmlns:d3p1="http://schemas.datacontract.org/2004/07/Microsoft.Office.Server.Search.Administration">
      <d3p1:LastItemName i:nil="true" />
      <d3p1:dictionary xmlns:d4p1="http://schemas.microsoft.com/2003/10/Serialization/Arrays">
        <d4p1:KeyValueOfstringOverrideInfoy6h3NzC8>
          <d4p1:Key>${ManagedPid}</d4p1:Key>
          <d4p1:Value>
            <d3p1:AliasesOverridden>false</d3p1:AliasesOverridden>
            <d3p1:EntityExtractorBitMap>0</d3p1:EntityExtractorBitMap>
            <d3p1:ExtraProperties i:nil="true" />
            <d3p1:ManagedPid>${ManagedPid}</d3p1:ManagedPid>
            <d3p1:MappingsOverridden>false</d3p1:MappingsOverridden>
            <d3p1:Name>${ManagedPid}</d3p1:Name>
            <d3p1:SchemaId>143692</d3p1:SchemaId>
            <d3p1:TokenNormalization>true</d3p1:TokenNormalization>
          </d4p1:Value>
        </d4p1:KeyValueOfstringOverrideInfoy6h3NzC8>
      </d3p1:dictionary>
    </Overrides>
  </SearchSchemaConfigurationSettings>
  <SearchSubscriptionSettingsConfigurationSettings i:nil="true" />
  <SearchTaxonomyConfigurationSettings i:nil="true" />
</SearchConfigurationSettings>
"@
}

# ═══════════════════════════════════════════════════════════════════════════
# Main Execution
# ═══════════════════════════════════════════════════════════════════════════

try {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host " SP Search — Crawled Property Mapping" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Site:     $SiteUrl" -ForegroundColor Gray
    Write-Host "  Mappings: $($mappings.Count)" -ForegroundColor Gray
    Write-Host ""

    # ─── Connect ──────────────────────────────────────────────────────────
    Write-Host "[1/3] Connecting to SharePoint..." -ForegroundColor Cyan
    $connectParams = @{ Url = $SiteUrl; Interactive = $true }
    if ($ClientId) { $connectParams.ClientId = $ClientId }
    Connect-PnPOnline @connectParams
    Write-Host "  Connected" -ForegroundColor Green
    Write-Host ""

    # ─── Apply Mappings ───────────────────────────────────────────────────
    Write-Host "[2/3] Applying managed property mappings..." -ForegroundColor Cyan
    Write-Host ""

    $succeeded = 0
    $failed = 0

    foreach ($m in $mappings) {
        $cp   = $m.CrawledProperty
        $mp   = $m.ManagedProperty
        $ps   = $m.Propset
        $ft   = $m.FilterType
        $mpid = Get-ManagedPid -ManagedProperty $mp

        Write-Host "  $cp -> $mp (PID: $mpid)" -ForegroundColor Yellow -NoNewline

        try {
            $xml = Get-MappingXml -CrawledPropertyName $cp -PropsetGuid $ps -ManagedPid $mpid
            Set-PnPSearchConfiguration -Configuration $xml -Scope Site
            Write-Host " [OK]" -ForegroundColor Green
            $succeeded++
        }
        catch {
            Write-Host " [FAILED]" -ForegroundColor Red
            Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }

        # Brief pause to avoid throttling
        Start-Sleep -Milliseconds 500
    }

    Write-Host ""

    # ─── Re-Index ─────────────────────────────────────────────────────────
    if ($RequestReindex) {
        Write-Host "[3/3] Requesting site re-index..." -ForegroundColor Cyan
        Request-PnPReIndexWeb
        Write-Host "  Re-index requested (15-60 minutes)" -ForegroundColor Green
    } else {
        Write-Host "[3/3] Skipping re-index (use -RequestReindex to enable)" -ForegroundColor Yellow
    }

    # ─── Summary ──────────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host " Mapping Complete!" -ForegroundColor Green
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "  Succeeded: $succeeded / $($mappings.Count)" -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
    if ($failed -gt 0) {
        Write-Host "  Failed:    $failed" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "  Mapping Summary:" -ForegroundColor White
    Write-Host "  ─────────────────────────────────────────────────────────" -ForegroundColor Gray
    Write-Host "  Managed Property     | Crawled Property          | Filter" -ForegroundColor White
    Write-Host "  ─────────────────────────────────────────────────────────" -ForegroundColor Gray
    foreach ($m in $mappings) {
        Write-Host "  $($m.ManagedProperty.PadRight(20)) | $($m.CrawledProperty.PadRight(25)) | $($m.FilterType)" -ForegroundColor Gray
    }
    Write-Host ""

    if (-not $RequestReindex) {
        Write-Host "  IMPORTANT: Run 'Request-PnPReIndexWeb' or pass -RequestReindex" -ForegroundColor Yellow
        Write-Host "  to trigger a re-crawl. Mappings take effect after the next crawl." -ForegroundColor Yellow
        Write-Host ""
    }

    Write-Host "  Verify mappings:" -ForegroundColor White
    Write-Host "  Get-PnPSearchConfiguration -Scope Site -OutputFormat ManagedPropertyMappings" -ForegroundColor DarkGray
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host " Mapping failed!" -ForegroundColor Red
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Stack: $($_.ScriptStackTrace)" -ForegroundColor DarkGray
    Write-Host ""
    exit 1
} finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    Write-Host "Disconnected from SharePoint" -ForegroundColor DarkGray
}
