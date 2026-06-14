<#
.SYNOPSIS
    Provisions an Account Documents search test environment end-to-end.

.DESCRIPTION
    Creates:
    - SP Search hidden support lists
    - Account taxonomy and supporting term sets
    - Site columns and 10 document content types
    - 6 document libraries with overlapping content types
    - Library-specific calculated subtype fields
    - Realistic office/PDF/TXT test documents with metadata
    - A modern search page with matching results, filters, verticals, and admin manager configuration

    This is intended to mirror a realistic "Account Documents" search estate
    more closely than the generic Provision-TestData.ps1 script.

.PARAMETER SiteUrl
    Target SharePoint site URL. Defaults to https://dodgeandcox.sharepoint.com/sites/SPSearch/

.PARAMETER ClientId
    Entra app client ID for PnP interactive auth. Defaults to 970bb320-0d49-4b4a-aa8f-c3f4b1e5928f

.PARAMETER PageName
    Search page file name without .aspx.

.PARAMETER PageTitle
    Search page title.

.PARAMETER SearchContextId
    Shared search context ID used by all search web parts.

.PARAMETER DocumentsPerLibrary
    Number of documents to create in each provisioned library.

.PARAMETER CleanExisting
    Reset known provisioned libraries/lists before provisioning.

.PARAMETER RequestReindex
    Request site reindex after provisioning.

.PARAMETER Publish
    Publish the search page after provisioning.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl = 'https://dodgeandcox.sharepoint.com/sites/SPSearch/',

    [Parameter(Mandatory = $false)]
    [string]$ClientId = '970bb320-0d49-4b4a-aa8f-c3f4b1e5928f',

    [Parameter(Mandatory = $false)]
    [string]$PageName = 'AccountDocumentsSearch',

    [Parameter(Mandatory = $false)]
    [string]$PageTitle = 'Account Documents Search',

    [Parameter(Mandatory = $false)]
    [string]$SearchContextId = 'account-documents',

    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 500)]
    [int]$DocumentsPerLibrary = 500,

    [Parameter(Mandatory = $false)]
    [switch]$CleanExisting,

    [Parameter(Mandatory = $false)]
    [switch]$RequestReindex,

    [Parameter(Mandatory = $false)]
    [bool]$Publish = $true
)

$ErrorActionPreference = 'Stop'

$WP_SEARCH_BOX = 'SP Search Box'
$WP_SEARCH_RESULTS = 'SP Search Results'
$WP_SEARCH_FILTERS = 'SP Search Filters'
$WP_SEARCH_VERTICALS = 'SP Search Verticals'
$WP_SEARCH_ADMIN_MANAGER = 'SP Search Admin Manager'

$TERM_GROUP_NAME = 'Global'
$SITE_COLUMN_GROUP = 'SP Search Account Documents'
$CONTENT_TYPE_GROUP = 'SP Search Account Documents'
$ACCOUNT_TERMSET_NAME = 'Accounts'
$TAG_TERMSET_NAME = 'Account Tags'

$script:stats = @{
    TermSetsCreated    = 0
    TermsCreated       = 0
    SiteColumnsCreated = 0
    ContentTypesCreated = 0
    LibrariesCreated   = 0
    FoldersCreated     = 0
    DocumentsUploaded  = 0
    Errors             = 0
}

$script:ResolvedUsers = @()

$script:AccountTags = @(
    'Urgent',
    'Quarter-End',
    'Regulatory',
    'Client Sensitive',
    'Renewal',
    'Exception',
    'Audit',
    'Inactive'
)

$script:JurisdictionChoices = @('US', 'UK', 'EU', 'Singapore', 'India', 'Australia')
$script:RegionChoices = @('North America', 'EMEA', 'APAC', 'LATAM')
$script:ReviewStatusChoices = @('Draft', 'In Review', 'Approved', 'Expired', 'Inactive')

$script:DocumentTypeDefs = @(
    @{
        Name = 'Account Agreement'
        Description = 'Master account agreements and amendments'
        SubTypeField = 'ADAccountAgreementSubType'
        SubTypes = @('Master Agreement', 'Fee Schedule', 'Terms and Conditions', 'Amendment', 'Renewal Notice')
        Fields = @('ADAccount', 'ADEffectiveDate', 'ADExpiryDate', 'ADInactive', 'ADCounterparty', 'ADReviewOwner', 'ADJurisdictions', 'ADMonitoringLink')
    },
    @{
        Name = 'Account Mandate'
        Description = 'Mandates and operating instructions'
        SubTypeField = 'ADAccountMandateSubType'
        SubTypes = @('Signatory Mandate', 'Payment Mandate', 'Trading Mandate', 'Treasury Mandate', 'Settlement Instruction')
        Fields = @('ADAccount', 'ADEffectiveDate', 'ADExpiryDate', 'ADInactive', 'ADRequiresWetSignature', 'ADReviewOwner', 'ADTags')
    },
    @{
        Name = 'Bank Valuation Sample'
        Description = 'Valuation and pricing reference documents'
        SubTypeField = 'ADBankValuationSampleSubType'
        SubTypes = @('Monthly Valuation', 'Quarterly Valuation', 'Pricing Snapshot', 'NAV Support', 'Holdings Extract')
        Fields = @('ADAccount', 'ADEffectiveDate', 'ADExpiryDate', 'ADInactive', 'ADDocumentAmount', 'ADRiskScore', 'ADRegion', 'ADOperationalNotes')
    },
    @{
        Name = 'Entity Formation'
        Description = 'Formation and constitutional entity records'
        SubTypeField = 'ADEntityFormationSubType'
        SubTypes = @('Certificate of Incorporation', 'Operating Agreement', 'Board Resolution', 'Shareholder Register', 'Power of Attorney')
        Fields = @('ADAccount', 'ADEffectiveDate', 'ADInactive', 'ADCounterparty', 'ADJurisdictions', 'ADMonitoringLink')
    },
    @{
        Name = 'Entity Compliance'
        Description = 'Entity compliance and KYC records'
        SubTypeField = 'ADEntityComplianceSubType'
        SubTypes = @('Annual Return', 'UBO Declaration', 'KYC Pack', 'Officer Register', 'Good Standing Certificate')
        Fields = @('ADAccount', 'ADEffectiveDate', 'ADExpiryDate', 'ADInactive', 'ADReviewStatus', 'ADReviewOwner', 'ADTags')
    },
    @{
        Name = 'Operational Procedure'
        Description = 'Operating procedures and playbooks'
        SubTypeField = 'ADOperationalProcedureSubType'
        SubTypes = @('Payment Procedure', 'Reconciliation Procedure', 'Exception Handling', 'End-of-Day Checklist', 'Access Control Procedure')
        Fields = @('ADAccount', 'ADEffectiveDate', 'ADInactive', 'ADReviewStatus', 'ADOperationalNotes', 'ADTags')
    },
    @{
        Name = 'Tax Certificate'
        Description = 'Tax forms and tax residency records'
        SubTypeField = 'ADTaxCertificateSubType'
        SubTypes = @('W-8BEN-E', 'W-9', 'CRS Self Certification', 'FATCA Form', 'Residency Certificate')
        Fields = @('ADAccount', 'ADEffectiveDate', 'ADExpiryDate', 'ADInactive', 'ADCounterparty', 'ADJurisdictions')
    },
    @{
        Name = 'Regulatory Filing'
        Description = 'Regulatory notices and submissions'
        SubTypeField = 'ADRegulatoryFilingSubType'
        SubTypes = @('AML Filing', 'Licensing Submission', 'Regulatory Notice', 'Periodic Return', 'Breach Notification')
        Fields = @('ADAccount', 'ADEffectiveDate', 'ADExpiryDate', 'ADInactive', 'ADReviewStatus', 'ADRegion', 'ADMonitoringLink')
    },
    @{
        Name = 'Audit Support'
        Description = 'Audit evidence and walkthrough materials'
        SubTypeField = 'ADAuditSupportSubType'
        SubTypes = @('Evidence Pack', 'Control Walkthrough', 'Sampling Extract', 'Management Response', 'Audit Request Log')
        Fields = @('ADAccount', 'ADEffectiveDate', 'ADExpiryDate', 'ADInactive', 'ADReviewOwner', 'ADOperationalNotes', 'ADTags')
    },
    @{
        Name = 'Account Closure'
        Description = 'Closure and archival records'
        SubTypeField = 'ADAccountClosureSubType'
        SubTypes = @('Closure Request', 'Closure Approval', 'Balance Confirmation', 'Final Statement', 'Archive Checklist')
        Fields = @('ADAccount', 'ADEffectiveDate', 'ADExpiryDate', 'ADInactive', 'ADRetentionYears', 'ADReviewStatus', 'ADReviewOwner')
    }
)

$script:LibraryDefs = @(
    @{
        Name = 'AccountDocs'
        Description = 'Account-facing operating and legal documents'
        ContentTypes = @('Account Agreement', 'Account Mandate', 'Tax Certificate', 'Account Closure')
        FolderPaths = @('Active/2025', 'Active/2026', 'Pending Renewal', 'Archive')
    },
    @{
        Name = 'EntityDocs'
        Description = 'Entity records and governance material'
        ContentTypes = @('Entity Formation', 'Entity Compliance', 'Regulatory Filing')
        FolderPaths = @('Constitutions', 'Governance', 'Annual Returns', 'Archive')
    },
    @{
        Name = 'OperationalDocs'
        Description = 'Operational procedures and runbooks'
        ContentTypes = @('Operational Procedure', 'Account Mandate', 'Audit Support')
        FolderPaths = @('Payments', 'Controls', 'Exceptions', 'Archive')
    },
    @{
        Name = 'ComplianceDocs'
        Description = 'Compliance, tax, and regulatory records'
        ContentTypes = @('Entity Compliance', 'Tax Certificate', 'Regulatory Filing', 'Audit Support')
        FolderPaths = @('KYC', 'Tax', 'Regulatory', 'Archive')
    },
    @{
        Name = 'ValuationDocs'
        Description = 'Valuation support, samples, and evidence'
        ContentTypes = @('Bank Valuation Sample', 'Audit Support', 'Account Agreement')
        FolderPaths = @('Monthly', 'Quarterly', 'Year End', 'Archive')
    },
    @{
        Name = 'ArchiveDocs'
        Description = 'Inactive and retained account document records'
        ContentTypes = @('Account Closure', 'Entity Compliance', 'Audit Support')
        FolderPaths = @('Inactive Accounts', 'Expired Certificates', 'Retention Hold')
    }
)

function Resolve-PnPClientId {
    param(
        [string]$ExplicitClientId
    )

    if ($ExplicitClientId) {
        return $ExplicitClientId
    }

    $candidateNames = @(
        'ENTRAID_APP_ID',
        'ENTRAID_CLIENT_ID',
        'AZURE_CLIENT_ID'
    )

    foreach ($candidateName in $candidateNames) {
        $candidateValue = [Environment]::GetEnvironmentVariable($candidateName)
        if (-not [string]::IsNullOrWhiteSpace($candidateValue)) {
            return $candidateValue.Trim()
        }
    }

    throw 'PnP interactive auth now requires an Entra app client ID. Re-run with -ClientId <app-id>, or set one of these environment variables before running: ENTRAID_APP_ID, ENTRAID_CLIENT_ID, AZURE_CLIENT_ID.'
}

function Normalize-SiteUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return $Value.Trim().TrimEnd('/').ToLowerInvariant()
}

function Use-PnPConnection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetSiteUrl,

        [string]$ExplicitClientId
    )

    $connectionState = @{
        DisconnectOnExit = $false
    }

    $existingConnection = $null
    try {
        $existingConnection = Get-PnPConnection -ErrorAction Stop
    } catch {
        $existingConnection = $null
    }

    $normalizedTargetSiteUrl = Normalize-SiteUrl -Value $TargetSiteUrl
    $currentConnectionUrl = ''
    if ($existingConnection -and $existingConnection.Url) {
        $currentConnectionUrl = Normalize-SiteUrl -Value $existingConnection.Url
    }

    if ($existingConnection -and $currentConnectionUrl -eq $normalizedTargetSiteUrl) {
        Write-Host '  Reusing existing PnP connection' -ForegroundColor Green
        return $connectionState
    }

    $resolvedClientId = Resolve-PnPClientId -ExplicitClientId $ExplicitClientId
    Connect-PnPOnline -Url $TargetSiteUrl -ClientId $resolvedClientId -Interactive
    $connectionState.DisconnectOnExit = $true
    Write-Host "  Connected to $TargetSiteUrl" -ForegroundColor Green
    return $connectionState
}

function Ensure-HiddenList {
    param(
        [string]$ListName,
        [string]$Description
    )

    $list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($list) {
        return $list
    }

    $list = New-PnPList -Title $ListName -Template GenericList -Hidden -EnableVersioning -OnQuickLaunch:$false
    Set-PnPList -Identity $ListName -Description $Description
    return $list
}

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
        return
    }

    switch ($FieldType) {
        'Text' {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type Text -Required:$Required | Out-Null
        }
        'Note' {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type Note -Required:$Required | Out-Null
        }
        'Number' {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type Number -Required:$Required | Out-Null
        }
        'DateTime' {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type DateTime -Required:$Required | Out-Null
        }
        'URL' {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type URL -Required:$Required | Out-Null
        }
        'Boolean' {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type Boolean -Required:$Required | Out-Null
        }
        'Choice' {
            Add-PnPField -List $ListName -DisplayName $FieldName -InternalName $FieldName -Type Choice -Choices $Choices -Required:$Required | Out-Null
        }
        'UserMulti' {
            Add-PnPFieldFromXml -List $ListName -FieldXml "<Field Type='UserMulti' DisplayName='$FieldName' StaticName='$FieldName' Name='$FieldName' Mult='TRUE' UserSelectionMode='PeopleOnly' />" | Out-Null
        }
    }
}

function Ensure-Index {
    param(
        [string]$ListName,
        [string]$FieldName
    )

    $field = Get-PnPField -List $ListName -Identity $FieldName -ErrorAction SilentlyContinue
    if (-not $field -or $field.Indexed) {
        return
    }

    $field.Indexed = $true
    $field.Update()
    Invoke-PnPQuery
}

function Ensure-SearchSupportLists {
    Write-Host '  Ensuring SP Search hidden lists...' -ForegroundColor Gray

    Ensure-HiddenList -ListName 'SearchSavedQueries' -Description 'SP Search: Saved/shared searches and state snapshots' | Out-Null
    Ensure-Field -ListName 'SearchSavedQueries' -FieldName 'QueryText' -FieldType 'Note'
    Ensure-Field -ListName 'SearchSavedQueries' -FieldName 'SearchState' -FieldType 'Note'
    Ensure-Field -ListName 'SearchSavedQueries' -FieldName 'SearchUrl' -FieldType 'URL'
    Ensure-Field -ListName 'SearchSavedQueries' -FieldName 'EntryType' -FieldType 'Choice' -Choices @('SavedSearch', 'SharedSearch', 'StateSnapshot')
    Ensure-Field -ListName 'SearchSavedQueries' -FieldName 'Category' -FieldType 'Text'
    Ensure-Field -ListName 'SearchSavedQueries' -FieldName 'SharedWith' -FieldType 'UserMulti'
    Ensure-Field -ListName 'SearchSavedQueries' -FieldName 'ResultCount' -FieldType 'Number'
    Ensure-Field -ListName 'SearchSavedQueries' -FieldName 'LastUsed' -FieldType 'DateTime'
    Ensure-Field -ListName 'SearchSavedQueries' -FieldName 'ExpiresAt' -FieldType 'DateTime'
    Ensure-Index -ListName 'SearchSavedQueries' -FieldName 'Author'
    Ensure-Index -ListName 'SearchSavedQueries' -FieldName 'Title'
    Ensure-Index -ListName 'SearchSavedQueries' -FieldName 'EntryType'
    Ensure-Index -ListName 'SearchSavedQueries' -FieldName 'Category'
    Ensure-Index -ListName 'SearchSavedQueries' -FieldName 'LastUsed'
    Ensure-Index -ListName 'SearchSavedQueries' -FieldName 'ExpiresAt'

    Ensure-HiddenList -ListName 'SearchHistory' -Description 'SP Search: User search history (high-volume, auto-pruned)' | Out-Null
    Ensure-Field -ListName 'SearchHistory' -FieldName 'QueryText' -FieldType 'Note'
    Ensure-Field -ListName 'SearchHistory' -FieldName 'QueryHash' -FieldType 'Text'
    Ensure-Field -ListName 'SearchHistory' -FieldName 'Vertical' -FieldType 'Text'
    Ensure-Field -ListName 'SearchHistory' -FieldName 'SearchPageUrl' -FieldType 'Text'
    Ensure-Field -ListName 'SearchHistory' -FieldName 'SearchState' -FieldType 'Note'
    Ensure-Field -ListName 'SearchHistory' -FieldName 'UseCount' -FieldType 'Number'
    Ensure-Field -ListName 'SearchHistory' -FieldName 'ResultCount' -FieldType 'Number'
    Ensure-Field -ListName 'SearchHistory' -FieldName 'IsZeroResult' -FieldType 'Boolean'
    Ensure-Field -ListName 'SearchHistory' -FieldName 'ClickedItems' -FieldType 'Note'
    Ensure-Field -ListName 'SearchHistory' -FieldName 'SearchTimestamp' -FieldType 'DateTime'
    Ensure-Index -ListName 'SearchHistory' -FieldName 'Author'
    Ensure-Index -ListName 'SearchHistory' -FieldName 'SearchTimestamp'
    Ensure-Index -ListName 'SearchHistory' -FieldName 'QueryHash'
    Ensure-Index -ListName 'SearchHistory' -FieldName 'Vertical'
    Set-PnPList -Identity 'SearchHistory' -BreakRoleInheritance -ClearSubScopes | Out-Null
    Set-PnPList -Identity 'SearchHistory' -ReadSecurity 2 -WriteSecurity 2 | Out-Null

    Ensure-HiddenList -ListName 'SearchCollections' -Description 'SP Search: User search result collections/pinboards' | Out-Null
    Ensure-Field -ListName 'SearchCollections' -FieldName 'ItemUrl' -FieldType 'URL'
    Ensure-Field -ListName 'SearchCollections' -FieldName 'ItemTitle' -FieldType 'Text'
    Ensure-Field -ListName 'SearchCollections' -FieldName 'ItemMetadata' -FieldType 'Note'
    Ensure-Field -ListName 'SearchCollections' -FieldName 'CollectionName' -FieldType 'Text'
    Ensure-Field -ListName 'SearchCollections' -FieldName 'Tags' -FieldType 'Note'
    Ensure-Field -ListName 'SearchCollections' -FieldName 'SharedWith' -FieldType 'UserMulti'
    Ensure-Field -ListName 'SearchCollections' -FieldName 'SortOrder' -FieldType 'Number'
    Ensure-Index -ListName 'SearchCollections' -FieldName 'Author'
    Ensure-Index -ListName 'SearchCollections' -FieldName 'Title'
    Ensure-Index -ListName 'SearchCollections' -FieldName 'CollectionName'

    Ensure-HiddenList -ListName 'SearchTelemetryConfig' -Description 'SP Search: Optional telemetry configuration (disabled by default)' | Out-Null
    Ensure-Field -ListName 'SearchTelemetryConfig' -FieldName 'IsEnabled' -FieldType 'Boolean'
    Ensure-Field -ListName 'SearchTelemetryConfig' -FieldName 'DestinationEndpoint' -FieldType 'Text'
    Ensure-Field -ListName 'SearchTelemetryConfig' -FieldName 'BatchIntervalSeconds' -FieldType 'Number'
    Ensure-Field -ListName 'SearchTelemetryConfig' -FieldName 'BatchSizeMax' -FieldType 'Number'
    Ensure-Field -ListName 'SearchTelemetryConfig' -FieldName 'PrivacyAcknowledgedBy' -FieldType 'Text'
    Ensure-Field -ListName 'SearchTelemetryConfig' -FieldName 'PrivacyAcknowledgedAt' -FieldType 'DateTime'
    $telemetryConfigRows = Get-PnPListItem -List 'SearchTelemetryConfig' -PageSize 1 -ErrorAction SilentlyContinue
    if (-not $telemetryConfigRows -or $telemetryConfigRows.Count -eq 0) {
        Add-PnPListItem -List 'SearchTelemetryConfig' -Values @{
            Title                = 'SP Search Telemetry Config (single row)'
            IsEnabled            = $false
            DestinationEndpoint  = ''
            BatchIntervalSeconds = 300
            BatchSizeMax         = 50
        } | Out-Null
    }

    Ensure-HiddenList -ListName 'SearchTelemetryOptIn' -Description 'SP Search: Optional per-user telemetry consent records' | Out-Null
    Ensure-Field -ListName 'SearchTelemetryOptIn' -FieldName 'ConsentTimestamp' -FieldType 'DateTime'
    Ensure-Field -ListName 'SearchTelemetryOptIn' -FieldName 'ConsentVersion' -FieldType 'Text'
    Ensure-Field -ListName 'SearchTelemetryOptIn' -FieldName 'AnonHash' -FieldType 'Text'
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 3
    )

    $attempt = 0
    do {
        $attempt++
        try {
            return & $ScriptBlock
        } catch {
            if ($attempt -ge $MaxAttempts) {
                throw
            }
            Start-Sleep -Seconds ([Math]::Pow(2, $attempt))
        }
    } while ($true)
}

function Ensure-TermGroup {
    param([string]$Name)
    $group = Get-PnPTermGroup -Identity $Name -ErrorAction SilentlyContinue
    if ($group) {
        return $group
    }
    $group = New-PnPTermGroup -Name $Name
    $script:stats.TermSetsCreated++
    return $group
}

function Ensure-TermSet {
    param([string]$Name, [string]$GroupName)
    $termSets = @()
    try {
        $termSets = @(Get-PnPTermSet -TermGroup $GroupName -ErrorAction Stop)
    } catch {
        $termSets = @()
    }

    $existing = $termSets | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
    if ($existing) {
        return $existing
    }

    $termSet = Invoke-WithRetry -MaxAttempts 3 -ScriptBlock {
        New-PnPTermSet -Name $Name -TermGroup $GroupName -Lcid 1033 -ErrorAction Stop
    }
    $script:stats.TermSetsCreated++
    return $termSet
}

function Ensure-Term {
    param(
        [string]$Name,
        $TermSet,
        $TermGroup,
        [string]$ParentTermId = $null
    )

    try {
        if ($ParentTermId) {
            $existing = Get-PnPTerm -Identity $Name -TermSet $TermSet -TermGroup $TermGroup -ParentTermId $ParentTermId -ErrorAction SilentlyContinue
        } else {
            $existing = Get-PnPTerm -Identity $Name -TermSet $TermSet -TermGroup $TermGroup -ErrorAction SilentlyContinue
        }
    } catch {
        $existing = $null
    }

    if ($existing) {
        return $existing
    }

    if ($ParentTermId) {
        $term = Invoke-WithRetry -MaxAttempts 3 -ScriptBlock {
            New-PnPTerm -Name $Name -TermSet $TermSet -TermGroup $TermGroup -Lcid 1033 -ParentTermId $ParentTermId -ErrorAction Stop
        }
    } else {
        $term = Invoke-WithRetry -MaxAttempts 3 -ScriptBlock {
            New-PnPTerm -Name $Name -TermSet $TermSet -TermGroup $TermGroup -Lcid 1033 -ErrorAction Stop
        }
    }

    $script:stats.TermsCreated++
    return $term
}

function Ensure-FlatTerms {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Names,

        [Parameter(Mandatory = $true)]
        $TermSet,

        [Parameter(Mandatory = $true)]
        $TermGroup,

        [Parameter(Mandatory = $true)]
        [string]$Activity
    )

    $existingTerms = @()
    try {
        $existingTerms = @(Get-PnPTerm -TermSet $TermSet -TermGroup $TermGroup -ErrorAction Stop)
    } catch {
        $existingTerms = @()
    }

    $existingNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($existingTerm in $existingTerms) {
        if ($existingTerm -and $existingTerm.Name) {
            $null = $existingNames.Add($existingTerm.Name)
        }
    }

    $missingNames = New-Object System.Collections.Generic.List[string]
    foreach ($name in $Names) {
        if (-not $existingNames.Contains($name)) {
            $missingNames.Add($name)
        }
    }

    if ($missingNames.Count -eq 0) {
        Write-Host "    ${Activity}: reusing existing terms ($($Names.Count) already present)" -ForegroundColor DarkGray
        return
    }

    $total = $missingNames.Count
    for ($i = 0; $i -lt $total; $i++) {
        $name = $missingNames[$i]
        $percent = if ($total -gt 0) { [Math]::Round((($i + 1) / $total) * 100, 0) } else { 100 }

        Write-Progress -Activity $Activity -Status "$($i + 1) of $total" -PercentComplete $percent

        Invoke-WithRetry -MaxAttempts 3 -ScriptBlock {
            New-PnPTerm -Name $name -TermSet $TermSet -TermGroup $TermGroup -Lcid 1033 -ErrorAction Stop
        } | Out-Null

        $null = $existingNames.Add($name)
        $script:stats.TermsCreated++

        if ((($i + 1) % 100) -eq 0 -or ($i + 1) -eq $total) {
            Write-Host "    ${Activity}: $($i + 1) / $total" -ForegroundColor DarkGray
        }
    }

    Write-Progress -Activity $Activity -Completed
}

function Ensure-SiteColumn {
    param(
        [string]$DisplayName,
        [string]$InternalName,
        [string]$Type,
        [string]$Group,
        [string[]]$Choices = @(),
        [hashtable]$ExtraParams = @{}
    )

    $field = Get-PnPField -Identity $InternalName -ErrorAction SilentlyContinue
    if ($field) {
        return $field
    }

    $params = @{
        DisplayName  = $DisplayName
        InternalName = $InternalName
        Type         = $Type
        Group        = $Group
    }
    if ($Choices.Count -gt 0) {
        $params['Choices'] = $Choices
    }
    foreach ($key in $ExtraParams.Keys) {
        $params[$key] = $ExtraParams[$key]
    }

    $field = Add-PnPField @params
    $script:stats.SiteColumnsCreated++
    return $field
}

function Ensure-TaxonomyColumn {
    param(
        [string]$DisplayName,
        [string]$InternalName,
        [string]$Group,
        [string]$TermSetPath,
        [switch]$MultiValue
    )

    $field = Get-PnPField -Identity $InternalName -ErrorAction SilentlyContinue
    if ($field) {
        return $field
    }

    $params = @{
        DisplayName  = $DisplayName
        InternalName = $InternalName
        Group        = $Group
        TermSetPath  = $TermSetPath
    }
    if ($MultiValue) {
        $params['MultiValue'] = $true
    }

    $field = Add-PnPTaxonomyField @params
    $script:stats.SiteColumnsCreated++
    return $field
}

function Get-DocumentTypeDefinition {
    param([string]$ContentTypeName)
    return $script:DocumentTypeDefs | Where-Object { $_.Name -eq $ContentTypeName } | Select-Object -First 1
}

function Ensure-ContentType {
    param(
        [string]$Name,
        [string]$Description,
        [string[]]$Fields
    )

    $contentType = Get-PnPContentType -Identity $Name -ErrorAction SilentlyContinue
    if (-not $contentType) {
        $parentContentType = Get-PnPContentType -Identity '0x0101'
        $contentType = Add-PnPContentType -Name $Name -Description $Description -Group $CONTENT_TYPE_GROUP -ParentContentType $parentContentType
        $script:stats.ContentTypesCreated++
    }

    foreach ($fieldName in $Fields) {
        try {
            Add-PnPFieldToContentType -ContentType $Name -Field $fieldName -ErrorAction SilentlyContinue | Out-Null
        } catch {
            if ($_.Exception.Message -notmatch 'already exists') {
                throw
            }
        }
    }

    return $contentType
}

function Ensure-Library {
    param(
        [string]$Name,
        [string]$Description
    )

    $created = $false
    $normalizedName = $Name.ToLowerInvariant()

    $list = Get-PnPList -Identity $Name -ErrorAction SilentlyContinue
    if (-not $list) {
        $lists = @(Get-PnPList -Includes RootFolder -ErrorAction SilentlyContinue)
        $list = $lists | Where-Object {
            $_.Title -eq $Name -or
            ($_.RootFolder -and $_.RootFolder.Name -eq $Name) -or
            ($_.RootFolder -and $_.RootFolder.ServerRelativeUrl.ToLowerInvariant().EndsWith('/' + $normalizedName))
        } | Select-Object -First 1
    }

    if (-not $list) {
        try {
            $list = New-PnPList -Title $Name -Url $Name -Template DocumentLibrary -EnableVersioning -OnQuickLaunch
            $created = $true
        } catch {
            if ($_.Exception.Message -notmatch 'already exists') {
                throw
            }

            $lists = @(Get-PnPList -Includes RootFolder -ErrorAction SilentlyContinue)
            $list = $lists | Where-Object {
                $_.Title -eq $Name -or
                ($_.RootFolder -and $_.RootFolder.Name -eq $Name) -or
                ($_.RootFolder -and $_.RootFolder.ServerRelativeUrl.ToLowerInvariant().EndsWith('/' + $normalizedName))
            } | Select-Object -First 1
        }
    }

    if (-not $list) {
        throw "Unable to resolve or create library '$Name'."
    }

    Set-PnPList -Identity $list -Description $Description -EnableContentTypes $true -EnableFolderCreation $true -OpenDocumentsMode Browser | Out-Null

    if ($created) {
        $script:stats.LibrariesCreated++
    }

    return $list
}

function Ensure-FolderPath {
    param(
        [string]$LibraryName,
        [string]$FolderPath
    )

    $segments = $FolderPath -split '/'
    $currentPath = $LibraryName
    foreach ($segment in $segments) {
        $targetPath = "$currentPath/$segment"
        try {
            $folder = Get-PnPFolder -Url $targetPath -ErrorAction SilentlyContinue
        } catch {
            $folder = $null
        }
        if (-not $folder -or -not $folder.Name) {
            Add-PnPFolder -Name $segment -Folder $currentPath | Out-Null
            $script:stats.FoldersCreated++
        }
        $currentPath = $targetPath
    }
}

function Ensure-LibraryContentTypes {
    param(
        [string]$LibraryName,
        [string[]]$ContentTypes
    )

    Set-PnPList -Identity $LibraryName -EnableContentTypes $true

    foreach ($contentTypeName in $ContentTypes) {
        try {
            Add-PnPContentTypeToList -List $LibraryName -ContentType $contentTypeName -ErrorAction SilentlyContinue | Out-Null
        } catch {
            if ($_.Exception.Message -notmatch 'already exists') {
                throw
            }
        }
    }

    if ($ContentTypes.Count -gt 0) {
        Set-PnPDefaultContentTypeToList -List $LibraryName -ContentType $ContentTypes[0]
    }

    try {
        $listContentTypes = Get-PnPContentType -List $LibraryName -ErrorAction SilentlyContinue
        $documentContentType = $listContentTypes | Where-Object { $_.Name -eq 'Document' } | Select-Object -First 1
        if ($documentContentType) {
            Remove-PnPContentTypeFromList -List $LibraryName -ContentType 'Document' -ErrorAction SilentlyContinue
        }
    } catch { }
}

function Build-CalculatedSubtypeFormula {
    param(
        [string[]]$FieldReferences
    )

    if (-not $FieldReferences -or $FieldReferences.Count -eq 0) {
        return '=""'
    }

    $formula = '[' + $FieldReferences[$FieldReferences.Count - 1] + ']'
    for ($i = $FieldReferences.Count - 2; $i -ge 0; $i--) {
        $fieldName = $FieldReferences[$i]
        $formula = 'IF([' + $fieldName + ']<>""' + ',[' + $fieldName + '],' + $formula + ')'
    }

    return '=' + $formula
}

function Ensure-LibraryCalculatedSubtypeField {
    param(
        [string]$LibraryName,
        [string[]]$SubTypeFields
    )

    $fieldReferences = New-Object System.Collections.Generic.List[string]
    foreach ($subTypeField in $SubTypeFields) {
        $listField = Get-PnPField -List $LibraryName -Identity $subTypeField -ErrorAction SilentlyContinue
        if (-not $listField) {
            try {
                Add-PnPField -List $LibraryName -Field $subTypeField -ErrorAction Stop | Out-Null
                $listField = Get-PnPField -List $LibraryName -Identity $subTypeField -ErrorAction SilentlyContinue
            } catch {
                Write-Warning "Could not attach subtype field '$subTypeField' to '$LibraryName': $($_.Exception.Message)"
            }
        }

        if ($listField) {
            $displayName = if ($listField.Title) { [string]$listField.Title } else { [string]$subTypeField }
            $null = $fieldReferences.Add($displayName)
        }
    }

    $formula = Build-CalculatedSubtypeFormula -FieldReferences ($fieldReferences.ToArray())
    $escapedFormula = Escape-Xml -Value $formula

    $existing = Get-PnPField -List $LibraryName -Identity 'ADCalculatedSubType' -ErrorAction SilentlyContinue
    if (-not $existing) {
        $fieldXml = "<Field Type='Calculated' DisplayName='Calculated Sub Type' StaticName='ADCalculatedSubType' Name='ADCalculatedSubType' ResultType='Text' Group='$SITE_COLUMN_GROUP'><Formula>$escapedFormula</Formula></Field>"
        Add-PnPFieldFromXml -List $LibraryName -FieldXml $fieldXml | Out-Null
        return
    }

    try {
        Set-PnPField -List $LibraryName -Identity 'ADCalculatedSubType' -Values @{ Formula = $formula } | Out-Null
    } catch {
        Write-Warning "Could not update Calculated Sub Type formula on '$LibraryName': $($_.Exception.Message)"
    }
}

function Get-LibraryExistingDocumentCount {
    param(
        [string]$LibraryName,
        [int]$ExpectedFolderCount
    )

    try {
        $list = Get-PnPList -Identity $LibraryName -Includes ItemCount -ErrorAction Stop
        $itemCount = if ($null -ne $list.ItemCount) { [int]$list.ItemCount } else { 0 }
        $documentCount = $itemCount - $ExpectedFolderCount
        if ($documentCount -lt 0) {
            return 0
        }
        return $documentCount
    } catch {
        Write-Warning "Could not determine existing document count for '$LibraryName': $($_.Exception.Message)"
        return 0
    }
}

function Get-AccountTerms {
    $accounts = New-Object System.Collections.Generic.List[string]
    for ($i = 1000; $i -le 1999; $i++) {
        $accounts.Add($i.ToString())
    }
    return $accounts.ToArray()
}

function Ensure-TermStore {
    Write-Host '  Provisioning term store...' -ForegroundColor Gray

    $termGroup = Ensure-TermGroup -Name $TERM_GROUP_NAME
    $accountTermSet = Ensure-TermSet -Name $ACCOUNT_TERMSET_NAME -GroupName $TERM_GROUP_NAME
    Ensure-FlatTerms -Names (Get-AccountTerms) -TermSet $accountTermSet -TermGroup $termGroup -Activity 'Seeding account terms'

    $tagTermSet = Ensure-TermSet -Name $TAG_TERMSET_NAME -GroupName $TERM_GROUP_NAME
    Ensure-FlatTerms -Names $script:AccountTags -TermSet $tagTermSet -TermGroup $termGroup -Activity 'Seeding tag terms'
}

function Ensure-SiteColumns {
    Write-Host '  Creating site columns...' -ForegroundColor Gray

    Ensure-TaxonomyColumn -DisplayName 'Account' -InternalName 'ADAccount' -Group $SITE_COLUMN_GROUP -TermSetPath "$TERM_GROUP_NAME|$ACCOUNT_TERMSET_NAME" | Out-Null
    Ensure-TaxonomyColumn -DisplayName 'Account Tags' -InternalName 'ADTags' -Group $SITE_COLUMN_GROUP -TermSetPath "$TERM_GROUP_NAME|$TAG_TERMSET_NAME" -MultiValue | Out-Null

    Ensure-SiteColumn -DisplayName 'Expiry Date' -InternalName 'ADExpiryDate' -Type 'DateTime' -Group $SITE_COLUMN_GROUP | Out-Null
    Ensure-SiteColumn -DisplayName 'Inactive' -InternalName 'ADInactive' -Type 'Boolean' -Group $SITE_COLUMN_GROUP | Out-Null
    Ensure-SiteColumn -DisplayName 'Effective Date' -InternalName 'ADEffectiveDate' -Type 'DateTime' -Group $SITE_COLUMN_GROUP | Out-Null
    Ensure-SiteColumn -DisplayName 'Review Status' -InternalName 'ADReviewStatus' -Type 'Choice' -Group $SITE_COLUMN_GROUP -Choices $script:ReviewStatusChoices | Out-Null
    Ensure-SiteColumn -DisplayName 'Review Owner' -InternalName 'ADReviewOwner' -Type 'User' -Group $SITE_COLUMN_GROUP | Out-Null
    Ensure-SiteColumn -DisplayName 'Counterparty' -InternalName 'ADCounterparty' -Type 'Text' -Group $SITE_COLUMN_GROUP | Out-Null
    Ensure-SiteColumn -DisplayName 'Operational Notes' -InternalName 'ADOperationalNotes' -Type 'Note' -Group $SITE_COLUMN_GROUP | Out-Null
    Ensure-SiteColumn -DisplayName 'Requires Wet Signature' -InternalName 'ADRequiresWetSignature' -Type 'Boolean' -Group $SITE_COLUMN_GROUP | Out-Null
    Ensure-SiteColumn -DisplayName 'Document Amount' -InternalName 'ADDocumentAmount' -Type 'Currency' -Group $SITE_COLUMN_GROUP | Out-Null
    Ensure-SiteColumn -DisplayName 'Risk Score' -InternalName 'ADRiskScore' -Type 'Number' -Group $SITE_COLUMN_GROUP | Out-Null
    Ensure-SiteColumn -DisplayName 'Monitoring Link' -InternalName 'ADMonitoringLink' -Type 'URL' -Group $SITE_COLUMN_GROUP | Out-Null
    Ensure-SiteColumn -DisplayName 'Retention Years' -InternalName 'ADRetentionYears' -Type 'Number' -Group $SITE_COLUMN_GROUP | Out-Null
    Ensure-SiteColumn -DisplayName 'Region' -InternalName 'ADRegion' -Type 'Choice' -Group $SITE_COLUMN_GROUP -Choices $script:RegionChoices | Out-Null
    Ensure-SiteColumn -DisplayName 'Jurisdictions' -InternalName 'ADJurisdictions' -Type 'MultiChoice' -Group $SITE_COLUMN_GROUP -Choices $script:JurisdictionChoices | Out-Null

    foreach ($docType in $script:DocumentTypeDefs) {
        Ensure-SiteColumn -DisplayName ($docType.Name + ' Sub Type') -InternalName $docType.SubTypeField -Type 'Choice' -Group $SITE_COLUMN_GROUP -Choices $docType.SubTypes | Out-Null
    }
}

function Ensure-DocumentContentTypes {
    Write-Host '  Creating content types...' -ForegroundColor Gray

    foreach ($docType in $script:DocumentTypeDefs) {
        $fields = @($docType.SubTypeField) + $docType.Fields
        Ensure-ContentType -Name $docType.Name -Description $docType.Description -Fields $fields | Out-Null
    }
}

function Get-RandomUser {
    if ($script:ResolvedUsers.Count -eq 0) {
        return $null
    }
    return $script:ResolvedUsers[(Get-Random -Minimum 0 -Maximum $script:ResolvedUsers.Count)]
}

function Get-RandomArraySlice {
    param(
        [string[]]$Values,
        [int]$Min = 1,
        [int]$Max = 2
    )

    $count = Get-Random -Minimum $Min -Maximum ($Max + 1)
    return $Values | Sort-Object { Get-Random } | Select-Object -First $count
}

function Get-RandomSubtype {
    param(
        [hashtable]$DocumentTypeDefinition
    )

    return $DocumentTypeDefinition.SubTypes[(Get-Random -Minimum 0 -Maximum $DocumentTypeDefinition.SubTypes.Count)]
}

function Get-RandomExtension {
    param([int]$Seed)
    $extensions = @('docx', 'pdf', 'xlsx', 'txt', 'pptx')
    return $extensions[$Seed % $extensions.Count]
}

function Get-SafeFileName {
    param(
        [string]$Title,
        [string]$Extension
    )

    $safe = $Title -replace '[\\/:*?"<>|#%&{}~]', ' '
    $safe = ($safe -replace '\s+', ' ').Trim()
    if ($safe.Length -gt 90) {
        $safe = $safe.Substring(0, 90).Trim()
    }
    return "$safe.$Extension"
}

function Escape-Xml {
    param([string]$Value)
    return [System.Security.SecurityElement]::Escape($Value)
}

function New-DocxBytes {
    param([string]$Title, [string]$BodyText)

    $memStream = [System.IO.MemoryStream]::new()
    $archive = [System.IO.Compression.ZipArchive]::new($memStream, [System.IO.Compression.ZipArchiveMode]::Create, $true)

    $ct = $archive.CreateEntry('[Content_Types].xml')
    $w = [System.IO.StreamWriter]::new($ct.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    $w.Close()

    $rels = $archive.CreateEntry('_rels/.rels')
    $w = [System.IO.StreamWriter]::new($rels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    $w.Close()

    $wordRels = $archive.CreateEntry('word/_rels/document.xml.rels')
    $w = [System.IO.StreamWriter]::new($wordRels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    $w.Close()

    $doc = $archive.CreateEntry('word/document.xml')
    $w = [System.IO.StreamWriter]::new($doc.Open())
    $titleXml = Escape-Xml -Value $Title
    $bodyXml = Escape-Xml -Value $BodyText.Replace("`n", ' ')
    $w.Write("<?xml version='1.0' encoding='UTF-8' standalone='yes'?><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:body><w:p><w:r><w:t>$titleXml</w:t></w:r></w:p><w:p><w:r><w:t>$bodyXml</w:t></w:r></w:p></w:body></w:document>")
    $w.Close()

    $archive.Dispose()
    $bytes = $memStream.ToArray()
    $memStream.Dispose()
    return $bytes
}

function New-XlsxBytes {
    param([string]$Title, [string]$BodyText)

    $memStream = [System.IO.MemoryStream]::new()
    $archive = [System.IO.Compression.ZipArchive]::new($memStream, [System.IO.Compression.ZipArchiveMode]::Create, $true)

    $ct = $archive.CreateEntry('[Content_Types].xml')
    $w = [System.IO.StreamWriter]::new($ct.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
    $w.Close()

    $rels = $archive.CreateEntry('_rels/.rels')
    $w = [System.IO.StreamWriter]::new($rels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    $w.Close()

    $wbRels = $archive.CreateEntry('xl/_rels/workbook.xml.rels')
    $w = [System.IO.StreamWriter]::new($wbRels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    $w.Close()

    $workbook = $archive.CreateEntry('xl/workbook.xml')
    $w = [System.IO.StreamWriter]::new($workbook.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>')
    $w.Close()

    $shared = $archive.CreateEntry('xl/sharedStrings.xml')
    $w = [System.IO.StreamWriter]::new($shared.Open())
    $titleXml = Escape-Xml -Value $Title
    $bodyXml = Escape-Xml -Value $BodyText.Replace("`n", ' ')
    $w.Write("<?xml version='1.0' encoding='UTF-8' standalone='yes'?><sst xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' count='2' uniqueCount='2'><si><t>$titleXml</t></si><si><t>$bodyXml</t></si></sst>")
    $w.Close()

    $styles = $archive.CreateEntry('xl/styles.xml')
    $w = [System.IO.StreamWriter]::new($styles.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>')
    $w.Close()

    $sheet = $archive.CreateEntry('xl/worksheets/sheet1.xml')
    $w = [System.IO.StreamWriter]::new($sheet.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row></sheetData></worksheet>')
    $w.Close()

    $archive.Dispose()
    $bytes = $memStream.ToArray()
    $memStream.Dispose()
    return $bytes
}

function New-PptxBytes {
    param([string]$Title, [string]$BodyText)

    $memStream = [System.IO.MemoryStream]::new()
    $archive = [System.IO.Compression.ZipArchive]::new($memStream, [System.IO.Compression.ZipArchiveMode]::Create, $true)

    $ct = $archive.CreateEntry('[Content_Types].xml')
    $w = [System.IO.StreamWriter]::new($ct.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>')
    $w.Close()

    $rels = $archive.CreateEntry('_rels/.rels')
    $w = [System.IO.StreamWriter]::new($rels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
    $w.Close()

    $presentation = $archive.CreateEntry('ppt/presentation.xml')
    $w = [System.IO.StreamWriter]::new($presentation.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>')
    $w.Close()

    $presentationRels = $archive.CreateEntry('ppt/_rels/presentation.xml.rels')
    $w = [System.IO.StreamWriter]::new($presentationRels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>')
    $w.Close()

    $layout = $archive.CreateEntry('ppt/slideLayouts/slideLayout1.xml')
    $w = [System.IO.StreamWriter]::new($layout.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>')
    $w.Close()

    $master = $archive.CreateEntry('ppt/slideMasters/slideMaster1.xml')
    $w = [System.IO.StreamWriter]::new($master.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>')
    $w.Close()

    $masterRels = $archive.CreateEntry('ppt/slideMasters/_rels/slideMaster1.xml.rels')
    $w = [System.IO.StreamWriter]::new($masterRels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
    $w.Close()

    $slide = $archive.CreateEntry('ppt/slides/slide1.xml')
    $w = [System.IO.StreamWriter]::new($slide.Open())
    $titleXml = Escape-Xml -Value $Title
    $bodyXml = Escape-Xml -Value $BodyText.Replace("`n", ' ')
    $w.Write("<?xml version='1.0' encoding='UTF-8' standalone='yes'?><p:sld xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id='1' name=''/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id='2' name='Title'/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>$titleXml</a:t></a:r></a:p><a:p><a:r><a:t>$bodyXml</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>")
    $w.Close()

    $archive.Dispose()
    $bytes = $memStream.ToArray()
    $memStream.Dispose()
    return $bytes
}

function New-PdfBytes {
    param([string]$Title, [string]$BodyText)

    $safeTitle = ($Title -replace '[\r\n]+', ' ')
    $safeBody = ($BodyText -replace '[\r\n]+', ' ')
    $content = "BT /F1 18 Tf 50 760 Td ($safeTitle) Tj 0 -24 Td /F1 11 Tf ($safeBody) Tj ET"
    $pdf = @"
%PDF-1.4
1 0 obj <</Type /Catalog /Pages 2 0 R>> endobj
2 0 obj <</Type /Pages /Kids [3 0 R] /Count 1>> endobj
3 0 obj <</Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources <</Font <</F1 5 0 R>>>> /Contents 4 0 R>> endobj
4 0 obj <</Length $($content.Length)>> stream
$content
endstream endobj
5 0 obj <</Type /Font /Subtype /Type1 /BaseFont /Helvetica>> endobj
xref
0 6
0000000000 65535 f 
0000000010 00000 n 
0000000060 00000 n 
0000000117 00000 n 
0000000244 00000 n 
0000000380 00000 n 
trailer <</Root 1 0 R /Size 6>>
startxref
455
%%EOF
"@
    return [System.Text.Encoding]::ASCII.GetBytes($pdf)
}

function New-DocumentBytes {
    param(
        [string]$Extension,
        [string]$Title,
        [string]$Body
    )

    switch ($Extension.ToLowerInvariant()) {
        'docx' { return New-DocxBytes -Title $Title -BodyText $Body }
        'xlsx' { return New-XlsxBytes -Title $Title -BodyText $Body }
        'pptx' { return New-PptxBytes -Title $Title -BodyText $Body }
        'pdf'  { return New-PdfBytes -Title $Title -BodyText $Body }
        default { return [System.Text.Encoding]::UTF8.GetBytes($Body) }
    }
}

function Get-LibraryFileServerRelativeUrl {
    param(
        [string]$FolderPath,
        [string]$FileName
    )

    $siteUri = [Uri]$SiteUrl
    $sitePath = $siteUri.AbsolutePath.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($sitePath)) {
        $sitePath = ''
    }

    return "$sitePath/$FolderPath/$FileName"
}

function Get-CounterpartyName {
    param([string]$AccountNumber)
    $prefixes = @('Northwind', 'Contoso', 'Fabrikam', 'Apex', 'Summit', 'Vertex', 'Harbor', 'Crescent')
    $lastDigit = [int]$AccountNumber.Substring($AccountNumber.Length - 1, 1)
    return $prefixes[$lastDigit % $prefixes.Count] + ' Financial'
}

function Build-DocumentMetadata {
    param(
        [hashtable]$LibraryDefinition,
        [hashtable]$DocumentTypeDefinition,
        [int]$Index
    )

    $accounts = Get-AccountTerms
    $account = $accounts[$Index % $accounts.Count]
    $subType = Get-RandomSubtype -DocumentTypeDefinition $DocumentTypeDefinition
    $effectiveDate = (Get-Date).AddDays(-1 * (Get-Random -Minimum 5 -Maximum 540))
    $expiryDate = $effectiveDate.AddDays((Get-Random -Minimum 90 -Maximum 730))
    $inactive = ($LibraryDefinition.Name -eq 'ArchiveDocs') -or (($Index % 11) -eq 0)
    $reviewOwner = Get-RandomUser
    $region = $script:RegionChoices[$Index % $script:RegionChoices.Count]
    $tags = Get-RandomArraySlice -Values $script:AccountTags -Min 1 -Max 2
    $jurisdictions = Get-RandomArraySlice -Values $script:JurisdictionChoices -Min 1 -Max 2
    $reviewStatus = if ($inactive) { 'Inactive' } elseif ($expiryDate -lt (Get-Date)) { 'Expired' } else { $script:ReviewStatusChoices[$Index % 3] }

    $values = @{
        Title = "$account - $($DocumentTypeDefinition.Name) - $subType"
    }

    foreach ($fieldName in $DocumentTypeDefinition.Fields) {
        switch ($fieldName) {
            'ADAccount' { $values[$fieldName] = "$TERM_GROUP_NAME|$ACCOUNT_TERMSET_NAME|$account" }
            'ADEffectiveDate' { $values[$fieldName] = $effectiveDate.ToString('MM/dd/yyyy HH:mm') }
            'ADExpiryDate' { $values[$fieldName] = $expiryDate.ToString('MM/dd/yyyy HH:mm') }
            'ADInactive' { $values[$fieldName] = $inactive }
            'ADCounterparty' { $values[$fieldName] = Get-CounterpartyName -AccountNumber $account }
            'ADReviewOwner' { if ($reviewOwner) { $values[$fieldName] = $reviewOwner } }
            'ADJurisdictions' { $values[$fieldName] = $jurisdictions }
            'ADMonitoringLink' { $values[$fieldName] = "https://monitoring.contoso.example/accounts/$account, Monitoring record" }
            'ADRequiresWetSignature' { $values[$fieldName] = (($Index % 4) -eq 0) }
            'ADTags' { $values[$fieldName] = $tags | ForEach-Object { "$TERM_GROUP_NAME|$TAG_TERMSET_NAME|$_" } }
            'ADDocumentAmount' { $values[$fieldName] = [Math]::Round((Get-Random -Minimum 5000 -Maximum 500000) + ($Index * 17.25), 2) }
            'ADRiskScore' { $values[$fieldName] = [Math]::Round((Get-Random -Minimum 1 -Maximum 10) + (($Index % 10) / 10), 1) }
            'ADRegion' { $values[$fieldName] = $region }
            'ADOperationalNotes' { $values[$fieldName] = "Generated test note for $($DocumentTypeDefinition.Name) in $($LibraryDefinition.Name). Account $account. Sub type $subType." }
            'ADReviewStatus' { $values[$fieldName] = $reviewStatus }
            'ADRetentionYears' { $values[$fieldName] = (Get-Random -Minimum 3 -Maximum 8) }
        }
    }

    $values[$DocumentTypeDefinition.SubTypeField] = $subType

    return @{
        Account = $account
        SubType = $subType
        EffectiveDate = $effectiveDate
        ExpiryDate = $expiryDate
        Inactive = $inactive
        Values = $values
    }
}

function Upload-LibraryDocuments {
    param(
        [hashtable]$LibraryDefinition,
        [int]$LibraryIndex,
        [int]$LibraryCount,
        [int]$ExistingDocumentCount,
        [int]$DocumentsToUpload,
        [int]$UploadOffset,
        [int]$TotalUploadsPlanned
    )

    if ($DocumentsToUpload -le 0) {
        Write-Host "    Skipping library $LibraryIndex/${LibraryCount}: $($LibraryDefinition.Name) already has $ExistingDocumentCount/$DocumentsPerLibrary docs" -ForegroundColor DarkGray
        return
    }

    $folders = @('') + $LibraryDefinition.FolderPaths
    $templateByKey = @{}
    Write-Host "    Uploading library $LibraryIndex/${LibraryCount}: $($LibraryDefinition.Name) ($ExistingDocumentCount existing, $DocumentsToUpload new)" -ForegroundColor DarkGray

    for ($i = 0; $i -lt $DocumentsToUpload; $i++) {
        $seedIndex = $ExistingDocumentCount + $i
        $currentDoc = $i + 1
        $overallCurrent = $UploadOffset + $currentDoc
        $overallTotal = if ($TotalUploadsPlanned -gt 0) { $TotalUploadsPlanned } else { $DocumentsToUpload }
        $overallPercent = if ($overallTotal -gt 0) { [Math]::Round(($overallCurrent / $overallTotal) * 100, 0) } else { 100 }

        Write-Progress -Activity 'Uploading account document test content' `
            -Status "Library $LibraryIndex/$LibraryCount - $($LibraryDefinition.Name) - Upload $currentDoc/$DocumentsToUpload" `
            -PercentComplete $overallPercent

        $docTypeName = $LibraryDefinition.ContentTypes[$seedIndex % $LibraryDefinition.ContentTypes.Count]
        $docType = Get-DocumentTypeDefinition -ContentTypeName $docTypeName
        $metadata = Build-DocumentMetadata -LibraryDefinition $LibraryDefinition -DocumentTypeDefinition $docType -Index $seedIndex
        $extension = Get-RandomExtension -Seed $seedIndex
        $title = $metadata.Values.Title
        $fileName = Get-SafeFileName -Title ($title + ' Rev ' + (($seedIndex % 4) + 1)) -Extension $extension
        $folderPath = $folders[$seedIndex % $folders.Count]
        $targetFolder = if ([string]::IsNullOrWhiteSpace($folderPath)) { $LibraryDefinition.Name } else { "$($LibraryDefinition.Name)/$folderPath" }
        $targetServerRelativeUrl = Get-LibraryFileServerRelativeUrl -FolderPath $targetFolder -FileName $fileName
        $templateKey = $docType.Name + '|' + $extension

        $body = @(
            "Account document test file",
            "Library: $($LibraryDefinition.Name)",
            "Type: $($docType.Name)",
            "Sub Type: $($metadata.SubType)",
            "Account: $($metadata.Account)"
        ) -join "`n"

        $stream = $null
        try {
            if ($templateByKey.ContainsKey($templateKey)) {
                Invoke-WithRetry -ScriptBlock {
                    Copy-PnPFile -SourceUrl $templateByKey[$templateKey] -TargetUrl $targetServerRelativeUrl -OverwriteIfAlreadyExists -Force -Confirm:$false -ErrorAction Stop | Out-Null
                } | Out-Null

                $copiedItem = Invoke-WithRetry -ScriptBlock {
                    Get-PnPFile -Url $targetServerRelativeUrl -AsListItem -ErrorAction Stop
                }

                Invoke-WithRetry -ScriptBlock {
                    Set-PnPListItem -List $LibraryDefinition.Name -Identity $copiedItem.Id -Values $metadata.Values -ErrorAction Stop | Out-Null
                } | Out-Null
            } else {
                $bytes = New-DocumentBytes -Extension $extension -Title $title -Body $body
                $stream = [System.IO.MemoryStream]::new($bytes)
                Invoke-WithRetry -ScriptBlock {
                    Add-PnPFile -FileName $fileName -Folder $targetFolder -Stream $stream -ContentType $docType.Name -Values $metadata.Values -ErrorAction Stop
                } | Out-Null
                $stream.Dispose()
                $templateByKey[$templateKey] = $targetServerRelativeUrl
            }
            $script:stats.DocumentsUploaded++

            if (($currentDoc % 10) -eq 0 -or $currentDoc -eq $DocumentsToUpload) {
                Write-Host "      Uploaded $currentDoc/$DocumentsToUpload in $($LibraryDefinition.Name) (overall $overallCurrent/$overallTotal)" -ForegroundColor DarkGray
            }

        } catch {
            if ($stream) {
                try { $stream.Dispose() } catch { }
            }
            $script:stats.Errors++
            Write-Warning "  Failed to upload '$fileName' to '$($LibraryDefinition.Name)': $($_.Exception.Message)"
        }
    }

    Write-Host "    Completed $($LibraryDefinition.Name): $DocumentsToUpload docs uploaded" -ForegroundColor Gray
}

function Get-AccountCoverageProfiles {
    param(
        [string]$BaseSiteUrl
    )

    $normalizedSiteUrl = $BaseSiteUrl.TrimEnd('/')
    $profiles = @()

    $profiles += @{
        title = 'All Account Document Libraries'
        description = 'Coverage profile for all provisioned account document libraries.'
        sourceUrls = (($script:LibraryDefs | ForEach-Object { "$normalizedSiteUrl/$($_.Name)" }) -join ', ')
        queryTemplate = '{searchTerms} IsDocument:1'
        includeFolders = $false
        trimDuplicates = $false
    }

    foreach ($library in $script:LibraryDefs) {
        $profiles += @{
            title = $library.Name
            description = $library.Description
            sourceUrls = "$normalizedSiteUrl/$($library.Name)"
            queryTemplate = '{searchTerms} IsDocument:1'
            includeFolders = $false
            trimDuplicates = $false
        }
    }

    return $profiles
}

function Get-PageServerRelativeUrl {
    param(
        [string]$BaseSiteUrl,
        [string]$TargetPageName
    )

    $siteUri = [Uri]$BaseSiteUrl
    $sitePath = $siteUri.AbsolutePath.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($sitePath)) {
        $sitePath = ''
    }

    return "$sitePath/SitePages/$TargetPageName.aspx"
}

function Remove-ExistingPage {
    param(
        [string]$BaseSiteUrl,
        [string]$TargetPageName
    )

    $pageServerRelativeUrl = Get-PageServerRelativeUrl -BaseSiteUrl $BaseSiteUrl -TargetPageName $TargetPageName

    try {
        Get-PnPFile -Url $pageServerRelativeUrl -AsListItem -ErrorAction Stop | Out-Null
    } catch {
        return
    }

    try {
        Undo-PnPFileCheckedOut -Url $pageServerRelativeUrl -ErrorAction Stop
    } catch { }

    try {
        Remove-PnPFile -ServerRelativeUrl $pageServerRelativeUrl -Force -Recycle -ErrorAction Stop
    } catch {
        Remove-PnPPage -Identity $TargetPageName -Force -ErrorAction SilentlyContinue
    }
}

function Add-SPSearchWebPart {
    param(
        [string]$Page,
        [string]$ComponentName,
        [int]$Section,
        [int]$Column,
        [int]$Order = 1,
        [hashtable]$Properties
    )

    Add-PnPPageWebPart -Page $Page `
        -Component $ComponentName `
        -Section $Section `
        -Column $Column `
        -Order $Order `
        -WebPartProperties $Properties `
        -ErrorAction Stop | Out-Null
}

function Provision-SearchPage {
    $pageUrl = "$($SiteUrl.TrimEnd('/'))/SitePages/$PageName.aspx"
    $siteScope = $SiteUrl.TrimEnd('/')

    $verticals = @(
        @{ uniqueId = 'vert-all'; key = 'all'; label = 'All'; iconName = 'Search'; queryTemplate = '{searchTerms} IsDocument:1'; sortOrder = 1; isLink = $false; openBehavior = 'currentTab' },
        @{ uniqueId = 'vert-account'; key = 'account'; label = 'Account Docs'; iconName = 'Page'; queryTemplate = "{searchTerms} IsDocument:1 Path:`"$siteScope/AccountDocs`""; sortOrder = 2; isLink = $false; openBehavior = 'currentTab' },
        @{ uniqueId = 'vert-entity'; key = 'entity'; label = 'Entity Docs'; iconName = 'Page'; queryTemplate = "{searchTerms} IsDocument:1 Path:`"$siteScope/EntityDocs`""; sortOrder = 3; isLink = $false; openBehavior = 'currentTab' },
        @{ uniqueId = 'vert-operational'; key = 'operational'; label = 'Operational Docs'; iconName = 'Page'; queryTemplate = "{searchTerms} IsDocument:1 Path:`"$siteScope/OperationalDocs`""; sortOrder = 4; isLink = $false; openBehavior = 'currentTab' },
        @{ uniqueId = 'vert-compliance'; key = 'compliance'; label = 'Compliance Docs'; iconName = 'Shield'; queryTemplate = "{searchTerms} IsDocument:1 Path:`"$siteScope/ComplianceDocs`""; sortOrder = 5; isLink = $false; openBehavior = 'currentTab' },
        @{ uniqueId = 'vert-valuation'; key = 'valuation'; label = 'Valuation Docs'; iconName = 'Financial'; queryTemplate = "{searchTerms} IsDocument:1 Path:`"$siteScope/ValuationDocs`""; sortOrder = 6; isLink = $false; openBehavior = 'currentTab' },
        @{ uniqueId = 'vert-archive'; key = 'archive'; label = 'Archive Docs'; iconName = 'Archive'; queryTemplate = "{searchTerms} IsDocument:1 Path:`"$siteScope/ArchiveDocs`""; sortOrder = 7; isLink = $false; openBehavior = 'currentTab' }
    )

    $filtersCollection = @(
        @{ uniqueId = 'filter-file'; managedProperty = 'FileType'; label = 'File Type'; displayName = 'File Type'; urlAlias = 'ft'; filterType = 'checkbox'; operator = 'OR'; maxValues = 10; defaultExpanded = $true; showCount = $true; sortBy = 'count'; sortDirection = 'desc'; multiValues = $true },
        @{ uniqueId = 'filter-doc-type'; managedProperty = 'ContentType'; label = 'Document Type'; displayName = 'Document Type'; urlAlias = 'dt'; filterType = 'dropdown'; operator = 'OR'; maxValues = 15; defaultExpanded = $true; showCount = $false; sortBy = 'name'; sortDirection = 'asc'; multiValues = $false },
        @{ uniqueId = 'filter-subtype'; managedProperty = 'RefinableString120'; label = 'Sub Type'; displayName = 'Sub Type'; urlAlias = 'st'; filterType = 'dropdown'; operator = 'OR'; maxValues = 20; defaultExpanded = $true; showCount = $false; sortBy = 'name'; sortDirection = 'asc'; multiValues = $false; dependsOn = 'ContentType'; showWhenParentHasValue = $true; hideZeroCountValues = $true; resetWhenParentChanges = $true },
        @{ uniqueId = 'filter-account'; managedProperty = 'RefinableString121'; label = 'Account'; displayName = 'Account'; urlAlias = 'ac'; filterType = 'taxonomy'; operator = 'OR'; maxValues = 20; defaultExpanded = $true; showCount = $true; sortBy = 'name'; sortDirection = 'asc'; multiValues = $true },
        @{ uniqueId = 'filter-expiry'; managedProperty = 'RefinableDate10'; label = 'Expiry Date'; displayName = 'Expiry Date'; urlAlias = 'ed'; filterType = 'daterange'; operator = 'AND'; maxValues = 10; defaultExpanded = $false; showCount = $true; sortBy = 'count'; sortDirection = 'desc'; multiValues = $false },
        @{ uniqueId = 'filter-active'; managedProperty = 'RefinableString122'; label = 'Is Active'; displayName = 'Is Active'; urlAlias = 'ia'; filterType = 'toggle'; operator = 'OR'; maxValues = 1; defaultExpanded = $false; showCount = $false; sortBy = 'name'; sortDirection = 'asc'; multiValues = $false; trueLabel = 'Active'; falseLabel = 'Inactive'; invertBoolean = $true }
    )

    $selectedPropertiesCollection = @(
        @{ uniqueId = 'sp-0'; property = 'Title'; alias = 'Title' },
        @{ uniqueId = 'sp-1'; property = 'ContentType'; alias = 'Document Type' },
        @{ uniqueId = 'sp-2'; property = 'RefinableString120'; alias = 'Sub Type' },
        @{ uniqueId = 'sp-3'; property = 'RefinableString121'; alias = 'Account' },
        @{ uniqueId = 'sp-4'; property = 'RefinableString123'; alias = 'Review Status' },
        @{ uniqueId = 'sp-5'; property = 'RefinableString125'; alias = 'Region' },
        @{ uniqueId = 'sp-6'; property = 'RefinableDate11'; alias = 'Effective Date' },
        @{ uniqueId = 'sp-7'; property = 'RefinableDate10'; alias = 'Expiry Date' },
        @{ uniqueId = 'sp-8'; property = 'RefinableString122'; alias = 'Inactive' },
        @{ uniqueId = 'sp-9'; property = 'FileType'; alias = 'Type' },
        @{ uniqueId = 'sp-10'; property = 'LastModifiedTime'; alias = 'Modified' },
        @{ uniqueId = 'sp-11'; property = 'Path'; alias = 'URL' }
    )

    $compactPropertiesCollection = @(
        @{ uniqueId = 'cp-0'; property = 'ContentType' },
        @{ uniqueId = 'cp-1'; property = 'RefinableString120' },
        @{ uniqueId = 'cp-2'; property = 'RefinableString121' },
        @{ uniqueId = 'cp-3'; property = 'RefinableDate10' },
        @{ uniqueId = 'cp-4'; property = 'FileType' }
    )

    $gridPropertiesCollection = @(
        @{ uniqueId = 'gp-0'; property = 'ContentType' },
        @{ uniqueId = 'gp-1'; property = 'RefinableString120' },
        @{ uniqueId = 'gp-2'; property = 'RefinableString121' },
        @{ uniqueId = 'gp-3'; property = 'RefinableString123' },
        @{ uniqueId = 'gp-4'; property = 'RefinableString125' },
        @{ uniqueId = 'gp-5'; property = 'RefinableDate10' },
        @{ uniqueId = 'gp-6'; property = 'RefinableString122' },
        @{ uniqueId = 'gp-7'; property = 'FileType' },
        @{ uniqueId = 'gp-8'; property = 'LastModifiedTime' }
    )

    $sortablePropertiesCollection = @(
        @{ uniqueId = 'sort-0'; property = 'LastModifiedTime'; label = 'Date Modified'; direction = 'Descending' },
        @{ uniqueId = 'sort-1'; property = 'Title'; label = 'Title'; direction = 'Ascending' },
        @{ uniqueId = 'sort-2'; property = 'RefinableDate10'; label = 'Expiry Date'; direction = 'Ascending' },
        @{ uniqueId = 'sort-3'; property = 'RefinableDate11'; label = 'Effective Date'; direction = 'Descending' },
        @{ uniqueId = 'sort-4'; property = 'ContentType'; label = 'Document Type'; direction = 'Ascending' },
        @{ uniqueId = 'sort-5'; property = 'RefinableString121'; label = 'Account'; direction = 'Ascending' }
    )

    $resultsRefinementFilters = @(
        @{ uniqueId = 'rf-0'; property = 'Path'; operator = 'Contains'; value = $siteScope }
    )

    Write-Host "  Creating search page '$PageName.aspx'..." -ForegroundColor Gray
    Remove-ExistingPage -BaseSiteUrl $SiteUrl -TargetPageName $PageName
    Add-PnPPage -Name $PageName -Title $PageTitle -LayoutType Article -HeaderLayoutType NoImage -CommentsEnabled:$false | Out-Null

    Add-PnPPageSection -Page $PageName -SectionTemplate OneColumn -Order 1 | Out-Null
    Add-PnPPageSection -Page $PageName -SectionTemplate OneColumn -Order 2 | Out-Null
    Add-PnPPageSection -Page $PageName -SectionTemplate TwoColumnLeft -Order 3 | Out-Null
    Add-PnPPageSection -Page $PageName -SectionTemplate OneColumn -Order 4 | Out-Null

    Add-SPSearchWebPart -Page $PageName -ComponentName $WP_SEARCH_BOX -Section 1 -Column 1 -Properties @{
        searchContextId     = $SearchContextId
        placeholder         = 'Search account documents...'
        debounceMs          = 300
        searchBehavior      = 'both'
        enableScopeSelector = $true
        enableSuggestions   = $true
        enableSearchManager = $true
        enableQueryBuilder  = $false
    }

    Add-SPSearchWebPart -Page $PageName -ComponentName $WP_SEARCH_VERTICALS -Section 2 -Column 1 -Properties @{
        searchContextId    = $SearchContextId
        verticalsCollection = $verticals
        defaultVertical    = 'all'
        showCounts         = $true
        hideEmptyVerticals = $false
        tabStyle           = 'underline'
    }

    Add-SPSearchWebPart -Page $PageName -ComponentName $WP_SEARCH_RESULTS -Section 3 -Column 1 -Properties @{
        searchContextId              = $SearchContextId
        queryTemplate                = '{searchTerms} IsDocument:1'
        searchScope                  = 'custom'
        searchScopePath              = $siteScope
        selectedPropertiesCollection = $selectedPropertiesCollection
        compactPropertiesCollection  = $compactPropertiesCollection
        gridPropertiesCollection     = $gridPropertiesCollection
        refinementFiltersCollection  = $resultsRefinementFilters
        sortablePropertiesCollection = $sortablePropertiesCollection
        defaultLayout                = 'list'
        showListLayout               = $true
        showCompactLayout            = $true
        showGridLayout               = $true
        showCardLayout               = $false
        showPeopleLayout             = $false
        showGalleryLayout            = $false
        showResultCount              = $true
        showSortDropdown             = $true
        pageSize                     = 20
        enableSelection              = $true
        trimDuplicates               = $false
    }

    Add-SPSearchWebPart -Page $PageName -ComponentName $WP_SEARCH_FILTERS -Section 3 -Column 2 -Properties @{
        searchContextId        = $SearchContextId
        filtersCollection      = $filtersCollection
        applyMode              = 'instant'
        operatorBetweenFilters = 'AND'
        showClearAll           = $true
        enableVisualFilterBuilder = $false
    }

    Add-SPSearchWebPart -Page $PageName -ComponentName $WP_SEARCH_ADMIN_MANAGER -Section 4 -Column 1 -Properties @{
        searchContextId             = $SearchContextId
        coverageSourcePageUrl       = $pageUrl
        mode                        = 'standalone'
        defaultTab                  = 'coverage'
        enableCoverage              = $true
        enableHealth                = $true
        enableInsights              = $true
        coverageProfilesCollection  = (Get-AccountCoverageProfiles -BaseSiteUrl $SiteUrl)
    }

    if ($Publish) {
        Set-PnPPage -Identity $PageName -Publish | Out-Null
    }
}

function Print-MappingGuide {
    Write-Host ''
    Write-Host 'Managed Property Mapping Guide' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'Map these crawled properties after the first crawl completes:' -ForegroundColor White
    Write-Host ''
    Write-Host '  Crawled Property            | Managed Property    | Used For' -ForegroundColor White
    Write-Host '  ----------------------------|---------------------|-----------------------------' -ForegroundColor Gray
    Write-Host '  ows_ADCalculatedSubType     | RefinableString120  | Sub Type refiner + display' -ForegroundColor Gray
    Write-Host '  ows_taxId_ADAccount         | RefinableString121  | Account refiner' -ForegroundColor Gray
    Write-Host '  ows_ADInactive              | RefinableString122  | Inactive refiner + display' -ForegroundColor Gray
    Write-Host '  ows_ADReviewStatus          | RefinableString123  | Review Status refiner + display' -ForegroundColor Gray
    Write-Host '  ows_ADCounterparty          | RefinableString124  | Counterparty refiner' -ForegroundColor Gray
    Write-Host '  ows_ADRegion                | RefinableString125  | Region refiner + display' -ForegroundColor Gray
    Write-Host '  ows_ADJurisdictions         | RefinableString126  | Jurisdictions refiner' -ForegroundColor Gray
    Write-Host '  ows_taxId_ADTags            | RefinableString127  | Account Tags refiner' -ForegroundColor Gray
    Write-Host '  ows_ADRequiresWetSignature  | RefinableString128  | Wet Signature refiner' -ForegroundColor Gray
    Write-Host '  ows_ADExpiryDate            | RefinableDate10     | Expiry Date refiner + sort' -ForegroundColor Gray
    Write-Host '  ows_ADEffectiveDate         | RefinableDate11     | Effective Date refiner + sort' -ForegroundColor Gray
    Write-Host ''
    Write-Host 'Built-in managed properties already used directly:' -ForegroundColor White
    Write-Host '  ContentType, FileType, LastModifiedTime, Path, Title' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'After mapping, request another reindex/crawl before validating refiners.' -ForegroundColor Yellow
}

Write-Host ''
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host ' Account Documents Environment Provisioning' -ForegroundColor Cyan
Write-Host '======================================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host "  Site:      $SiteUrl" -ForegroundColor White
Write-Host "  Page:      $PageName.aspx" -ForegroundColor White
Write-Host "  Context:   $SearchContextId" -ForegroundColor White
Write-Host "  Docs/lib:  $DocumentsPerLibrary" -ForegroundColor White
Write-Host "  TotalDocs: $($DocumentsPerLibrary * $script:LibraryDefs.Count)" -ForegroundColor White
Write-Host ''

$connectionState = $null

try {
    Import-Module PnP.PowerShell -ErrorAction Stop
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop

    $connectionState = Use-PnPConnection -TargetSiteUrl $SiteUrl -ExplicitClientId $ClientId

    try {
        $currentUser = Get-PnPProperty -ClientObject (Get-PnPWeb) -Property CurrentUser
        $script:ResolvedUsers = @($currentUser.LoginName)
    } catch {
        $script:ResolvedUsers = @()
    }

    if ($CleanExisting) {
        $resetScriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'Reset-AccountDocumentsEnvironment.ps1'
        if (-not (Test-Path $resetScriptPath)) {
            throw "Reset script not found at $resetScriptPath"
        }

        Write-Host '  Clearing existing provisioned artifacts before provisioning...' -ForegroundColor Yellow
        $resetParams = @{
            SiteUrl = $SiteUrl
        }
        if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
            $resetParams['ClientId'] = $ClientId
        }

        & $resetScriptPath @resetParams
    }

    Ensure-SearchSupportLists
    Ensure-TermStore
    Ensure-SiteColumns
    Ensure-DocumentContentTypes

    Write-Host '  Creating libraries and folder structure...' -ForegroundColor Gray
    foreach ($library in $script:LibraryDefs) {
        Ensure-Library -Name $library.Name -Description $library.Description | Out-Null
        Ensure-LibraryContentTypes -LibraryName $library.Name -ContentTypes $library.ContentTypes
        foreach ($folderPath in $library.FolderPaths) {
            Ensure-FolderPath -LibraryName $library.Name -FolderPath $folderPath
        }

        $subTypeFields = @()
        foreach ($contentTypeName in $library.ContentTypes) {
            $definition = Get-DocumentTypeDefinition -ContentTypeName $contentTypeName
            if ($definition) {
                $subTypeFields += $definition.SubTypeField
            }
        }
        Ensure-LibraryCalculatedSubtypeField -LibraryName $library.Name -SubTypeFields $subTypeFields
    }

    Write-Host '  Uploading account document test content...' -ForegroundColor Gray
    $libraryCount = $script:LibraryDefs.Count
    $uploadPlans = @()
    $totalUploadsPlanned = 0

    for ($libraryIndex = 0; $libraryIndex -lt $libraryCount; $libraryIndex++) {
        $library = $script:LibraryDefs[$libraryIndex]
        $existingDocumentCount = Get-LibraryExistingDocumentCount -LibraryName $library.Name -ExpectedFolderCount $library.FolderPaths.Count
        $documentsToUpload = [Math]::Max($DocumentsPerLibrary - $existingDocumentCount, 0)
        $uploadPlans += @{
            LibraryDefinition = $library
            ExistingDocumentCount = $existingDocumentCount
            DocumentsToUpload = $documentsToUpload
        }
        $totalUploadsPlanned += $documentsToUpload
    }

    if ($totalUploadsPlanned -eq 0) {
        Write-Host "    All libraries already meet the target count ($DocumentsPerLibrary docs/lib). Skipping uploads." -ForegroundColor DarkGray
    } else {
        $uploadOffset = 0
        for ($libraryIndex = 0; $libraryIndex -lt $libraryCount; $libraryIndex++) {
            $plan = $uploadPlans[$libraryIndex]
            Upload-LibraryDocuments `
                -LibraryDefinition $plan.LibraryDefinition `
                -LibraryIndex ($libraryIndex + 1) `
                -LibraryCount $libraryCount `
                -ExistingDocumentCount $plan.ExistingDocumentCount `
                -DocumentsToUpload $plan.DocumentsToUpload `
                -UploadOffset $uploadOffset `
                -TotalUploadsPlanned $totalUploadsPlanned
            $uploadOffset += $plan.DocumentsToUpload
        }
    }
    Write-Progress -Activity 'Uploading account document test content' -Completed

    Provision-SearchPage

    if ($RequestReindex) {
        try {
            Request-PnPReindexWeb -ErrorAction Stop
            Write-Host '  Site reindex requested.' -ForegroundColor Yellow
        } catch {
            Write-Warning "Could not request site reindex: $($_.Exception.Message)"
        }
    }

    Print-MappingGuide

    Write-Host ''
    Write-Host '======================================================================' -ForegroundColor Green
    Write-Host ' Account Documents environment ready' -ForegroundColor Green
    Write-Host '======================================================================' -ForegroundColor Green
    Write-Host ''
    Write-Host "  Search page: $($SiteUrl.TrimEnd('/'))/SitePages/$PageName.aspx" -ForegroundColor White
    Write-Host "  Libraries:   $($script:LibraryDefs.Count)" -ForegroundColor White
    Write-Host "  ContentTypes:$($script:DocumentTypeDefs.Count)" -ForegroundColor White
    Write-Host "  Documents:   $($script:stats.DocumentsUploaded)" -ForegroundColor White
    Write-Host "  Errors:      $($script:stats.Errors)" -ForegroundColor White
    Write-Host ''
} catch {
    Write-Host ''
    Write-Host '======================================================================' -ForegroundColor Red
    Write-Host ' Account Documents provisioning failed' -ForegroundColor Red
    Write-Host '======================================================================' -ForegroundColor Red
    Write-Host ''
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host "Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    Write-Host ''
    Write-Host 'Stack:' -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    Write-Host ''
    exit 1
} finally {
    if ($connectionState -and $connectionState.DisconnectOnExit) {
        try {
            $null = Get-PnPConnection -ErrorAction Stop
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Write-Host 'Disconnected from SharePoint' -ForegroundColor Gray
        } catch { }
    }
}
