<#
.SYNOPSIS
    Provisions test data for the SP Search solution.

.DESCRIPTION
    Creates 10 document libraries, 10 custom lists, term sets, site columns,
    and uploads 1,200+ documents with varied metadata for comprehensive search testing.

.PARAMETER SiteUrl
    Target SharePoint site URL (e.g. https://pixelboy.sharepoint.com/sites/SPSearch)

.PARAMETER ClientId
    Azure AD app registration Client ID (optional, enables non-interactive auth)

.PARAMETER Users
    Array of user login names for distributing authorship metadata.
    If empty, the connected user is used for all items.

.PARAMETER DocumentsPerLibrary
    Number of documents to create per library (default: 120, total ~1,200)

.PARAMETER ItemsPerList
    Number of items to create per list (default: 60, total ~600)

.PARAMETER TermGroupName
    Name of the term group in the Term Store (default: "SP Search Test Data")

.PARAMETER SkipTermStore
    Skip term store provisioning

.PARAMETER SkipDocuments
    Skip document generation and upload

.PARAMETER SkipListItems
    Skip list item creation

.PARAMETER RequestReindex
    Request site re-index after provisioning

.PARAMETER CleanExisting
    Remove existing test libraries/lists before recreation

.PARAMETER SiteColumnGroup
    Site column group name (default: "SP Search Test")

.EXAMPLE
    .\Provision-TestData.ps1 -SiteUrl "https://pixelboy.sharepoint.com/sites/SPSearch" `
        -Users @("user1@pixelboy.onmicrosoft.com","user2@pixelboy.onmicrosoft.com") `
        -RequestReindex
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Target SharePoint site URL")]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string[]]$Users = @(),

    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 500)]
    [int]$DocumentsPerLibrary = 120,

    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 200)]
    [int]$ItemsPerList = 60,

    [Parameter(Mandatory = $false)]
    [string]$TermGroupName = "SP Search Test Data",

    [Parameter(Mandatory = $false)]
    [switch]$SkipTermStore,

    [Parameter(Mandatory = $false)]
    [switch]$SkipDocuments,

    [Parameter(Mandatory = $false)]
    [switch]$SkipListItems,

    [Parameter(Mandatory = $false)]
    [switch]$RequestReindex,

    [Parameter(Mandatory = $false)]
    [switch]$CleanExisting,

    [Parameter(Mandatory = $false)]
    [string]$SiteColumnGroup = "SP Search Test"
)

$ErrorActionPreference = "Stop"

# ─── Statistics ───────────────────────────────────────────────────────────────

$script:stats = @{
    TermSetsCreated    = 0
    TermsCreated       = 0
    SiteColumnsCreated = 0
    LibrariesCreated   = 0
    ListsCreated       = 0
    FoldersCreated     = 0
    DocumentsUploaded  = 0
    ListItemsCreated   = 0
    Errors             = 0
}

$totalSteps = 9
$step = 0

# ─── Configuration Data ──────────────────────────────────────────────────────

$script:DepartmentTerms = [ordered]@{
    "Executive"       = @("CEO Office", "Board Relations")
    "Sales"           = @("Inside Sales", "Enterprise Sales", "Channel Partners")
    "Marketing"       = @("Digital Marketing", "Brand & Communications", "Events")
    "Human Resources" = @("Talent Acquisition", "Learning & Development", "Employee Relations")
    "Finance"         = @("Accounting", "Treasury", "Financial Planning")
    "Engineering"     = @("Frontend", "Backend", "DevOps", "QA")
    "Legal"           = @("Corporate Law", "Compliance", "Intellectual Property")
    "Operations"      = @("Supply Chain", "Facilities", "IT Infrastructure")
}

$script:DocumentTypeTerms = @(
    "Report", "Policy", "Procedure", "Proposal", "Contract",
    "Presentation", "Memo", "Specification", "Guide", "Template"
)

$script:RegionTerms = [ordered]@{
    "North America" = @("United States", "Canada", "Mexico")
    "Europe"        = @("United Kingdom", "Germany", "France")
    "Asia Pacific"  = @("Japan", "Australia", "India")
    "Latin America" = @("Brazil", "Argentina", "Colombia")
}

$script:TagTerms = @(
    "Confidential", "Draft", "Final", "Urgent", "Board Review", "Quarterly",
    "Annual", "External", "Internal", "Template", "Archive", "Featured"
)

$script:StatusChoices = @("Draft", "In Review", "Approved", "Published", "Archived")
$script:PriorityChoices = @("Critical", "High", "Medium", "Low")
$script:RegionChoices = @("North America", "Europe", "Asia Pacific", "Latin America")

# Weighted distributions (cumulative thresholds for random assignment)
$script:StatusWeights = @(0.20, 0.35, 0.60, 0.90, 1.00)   # Draft 20%, In Review 15%, Approved 25%, Published 30%, Archived 10%
$script:PriorityWeights = @(0.10, 0.35, 0.75, 1.00)         # Critical 10%, High 25%, Medium 40%, Low 25%
$script:RegionWeights = @(0.35, 0.65, 0.85, 1.00)           # NA 35%, Europe 30%, APAC 20%, LATAM 15%
$script:DeptWeights = @(0.10, 0.25, 0.37, 0.47, 0.62, 0.80, 0.88, 1.00)  # Exec 10%, Sales 15%, Mktg 12%, HR 10%, Fin 15%, Eng 18%, Legal 8%, Ops 12%
$script:DeptNames = @("Executive", "Sales", "Marketing", "Human Resources", "Finance", "Engineering", "Legal", "Operations")

# Library definitions: Name, Description, Columns, FolderDepth, Themes
$script:LibraryDefs = @(
    @{ Name = "CorporatePolicies"; Desc = "Policies, procedures, guidelines"; Cols = @("SPSStatus","SPSDepartment","SPSDocumentType","SPSIsPublished","SPSReviewDate","SPSOwner"); FolderDepth = 3; Themes = @("Legal","Operations") }
    @{ Name = "SalesMaterials"; Desc = "Proposals, contracts, presentations"; Cols = @("SPSStatus","SPSDepartment","SPSRegion","SPSBudget","SPSTags","SPSIsActive"); FolderDepth = 2; Themes = @("Sales","Finance") }
    @{ Name = "MarketingContent"; Desc = "Campaigns, brand assets, collateral"; Cols = @("SPSStatus","SPSDepartment","SPSRegion","SPSTags","SPSRating"); FolderDepth = 2; Themes = @("Marketing","Sales") }
    @{ Name = "HRResources"; Desc = "Handbooks, forms, training materials"; Cols = @("SPSStatus","SPSDepartment","SPSDocumentType","SPSIsActive","SPSOwner"); FolderDepth = 3; Themes = @("HR","Operations") }
    @{ Name = "FinanceReports"; Desc = "Budgets, forecasts, audits"; Cols = @("SPSStatus","SPSDepartment","SPSBudget","SPSIsPublished","SPSExternalRef"); FolderDepth = 2; Themes = @("Finance","Projects") }
    @{ Name = "EngineeringDocs"; Desc = "Specs, designs, architecture docs"; Cols = @("SPSStatus","SPSDepartment","SPSPriority","SPSTags","SPSIsActive"); FolderDepth = 3; Themes = @("Engineering","Projects") }
    @{ Name = "LegalDocuments"; Desc = "Contracts, NDAs, compliance docs"; Cols = @("SPSStatus","SPSDepartment","SPSDocumentType","SPSIsPublished","SPSReviewDate"); FolderDepth = 2; Themes = @("Legal","Finance") }
    @{ Name = "ProjectFiles"; Desc = "Plans, status reports, deliverables"; Cols = @("SPSStatus","SPSDepartment","SPSPriority","SPSBudget","SPSRegion","SPSTags"); FolderDepth = 3; Themes = @("Projects","Engineering") }
    @{ Name = "MediaAssets"; Desc = "Images, presentations, infographics"; Cols = @("SPSDepartment","SPSTags","SPSRating","SPSIsActive"); FolderDepth = 1; Themes = @("Marketing","HR") }
    @{ Name = "KnowledgeBase"; Desc = "FAQs, how-tos, tutorials, guides"; Cols = @("SPSStatus","SPSDepartment","SPSDocumentType","SPSTags","SPSViewCount"); FolderDepth = 2; Themes = @("Engineering","Operations") }
)

# List definitions: Name, Description, ExtraColumns (list-specific beyond shared site columns)
$script:ListDefs = @(
    @{ Name = "Projects"; Desc = "Project tracking"; SharedCols = @("SPSStatus","SPSOwner","SPSBudget","SPSDepartment","SPSRegion","SPSPriority"); ExtraCols = @(
        @{ Name = "StartDate"; Type = "DateTime" }, @{ Name = "EndDate"; Type = "DateTime" }, @{ Name = "PercentComplete"; Type = "Number" }
    )}
    @{ Name = "Contacts"; Desc = "People directory"; SharedCols = @("SPSDepartment","SPSRegion","SPSIsActive"); ExtraCols = @(
        @{ Name = "Email"; Type = "Text" }, @{ Name = "Phone"; Type = "Text" }, @{ Name = "JobTitle"; Type = "Text" }
    )}
    @{ Name = "Tasks"; Desc = "Task tracking"; SharedCols = @("SPSStatus","SPSPriority","SPSDepartment"); ExtraCols = @(
        @{ Name = "DueDate"; Type = "DateTime" }, @{ Name = "Assignee"; Type = "User" }, @{ Name = "EstimatedHours"; Type = "Number" }
    )}
    @{ Name = "Events"; Desc = "Events calendar"; SharedCols = @("SPSDepartment"); ExtraCols = @(
        @{ Name = "EventDate"; Type = "DateTime" }, @{ Name = "EventEndDate"; Type = "DateTime" }, @{ Name = "Location"; Type = "Text" },
        @{ Name = "Organizer"; Type = "User" },
        @{ Name = "EventCategory"; Type = "Choice"; Choices = @("Meeting","Conference","Training","Social","Webinar") }
    )}
    @{ Name = "Inventory"; Desc = "Asset tracking"; SharedCols = @("SPSIsActive"); ExtraCols = @(
        @{ Name = "Quantity"; Type = "Number" }, @{ Name = "UnitPrice"; Type = "Currency" }, @{ Name = "Location"; Type = "Text" },
        @{ Name = "ItemCategory"; Type = "Choice"; Choices = @("Hardware","Software","Furniture","Supplies","Equipment") }
    )}
    @{ Name = "Announcements"; Desc = "News items"; SharedCols = @("SPSPriority"); ExtraCols = @(
        @{ Name = "Body"; Type = "Note" }, @{ Name = "PublishDate"; Type = "DateTime" }, @{ Name = "ExpiryDate"; Type = "DateTime" },
        @{ Name = "AnnouncementCategory"; Type = "Choice"; Choices = @("Company","Team","IT","HR","Facilities") }
    )}
    @{ Name = "Issues"; Desc = "Issue tracking"; SharedCols = @("SPSStatus","SPSDepartment"); ExtraCols = @(
        @{ Name = "Resolution"; Type = "Note" }, @{ Name = "ReportedDate"; Type = "DateTime" }, @{ Name = "Assignee"; Type = "User" },
        @{ Name = "Severity"; Type = "Choice"; Choices = @("Critical","Major","Minor","Trivial") }
    )}
    @{ Name = "FAQ"; Desc = "Questions and answers"; SharedCols = @("SPSTags"); ExtraCols = @(
        @{ Name = "Question"; Type = "Note" }, @{ Name = "Answer"; Type = "Note" }, @{ Name = "HelpfulCount"; Type = "Number" },
        @{ Name = "FAQCategory"; Type = "Choice"; Choices = @("General","Technical","HR","Finance","IT") }
    )}
    @{ Name = "Policies"; Desc = "Policy registry"; SharedCols = @("SPSReviewDate","SPSOwner","SPSDepartment","SPSIsPublished"); ExtraCols = @(
        @{ Name = "PolicyNumber"; Type = "Text" }, @{ Name = "PolicyVersion"; Type = "Text" }
    )}
    @{ Name = "Glossary"; Desc = "Terms and definitions"; SharedCols = @("SPSIsActive"); ExtraCols = @(
        @{ Name = "Definition"; Type = "Note" }, @{ Name = "RelatedTerms"; Type = "Text" },
        @{ Name = "GlossaryCategory"; Type = "Choice"; Choices = @("Technical","Business","Legal","Finance","HR") }
    )}
)

# ─── Content Themes ──────────────────────────────────────────────────────────

$script:ThemeData = @{
    "Finance" = @{
        Keywords = @("budget", "revenue", "quarterly", "forecast", "fiscal", "audit", "ROI", "capital", "expense", "profit")
        Phrase = "annual budget report"
        Titles = @(
            "Q{0} Revenue Analysis Report", "Annual Budget Proposal {0}", "Fiscal Year {0} Audit Summary",
            "Capital Expenditure Review {0}", "Quarterly Forecast Update Q{0}", "Expense Report Summary {0}",
            "ROI Analysis for Project {0}", "Profit Margin Assessment {0}", "Financial Planning Guide {0}",
            "Budget Variance Report Q{0}", "Revenue Projection Model {0}", "Cost Optimization Plan {0}"
        )
        Paragraphs = @(
            "This annual budget report provides a comprehensive overview of the fiscal year financial performance. Revenue targets were analyzed across all business units to identify growth opportunities and areas requiring cost optimization.",
            "The quarterly forecast update reflects current market conditions and adjusted projections for the remaining fiscal quarters. Capital expenditure has been reallocated to prioritize high-ROI initiatives.",
            "Our audit findings indicate strong financial controls across departments. The expense management process has yielded significant savings, improving overall profit margins by approximately twelve percent.",
            "Budget allocation for the upcoming fiscal year emphasizes digital transformation investments. Revenue diversification strategies include expanding into new market segments and optimizing existing product lines.",
            "The financial planning committee reviewed all capital requests and approved funding for strategic initiatives. ROI projections exceed the corporate hurdle rate for all approved projects."
        )
    }
    "HR" = @{
        Keywords = @("employee", "onboarding", "benefits", "performance", "handbook", "training", "diversity", "recruitment", "compensation", "retention")
        Phrase = "employee performance review"
        Titles = @(
            "Employee Handbook {0} Edition", "Performance Review Guidelines {0}", "Benefits Enrollment Guide {0}",
            "Onboarding Checklist for {0}", "Training Program Catalog {0}", "Diversity and Inclusion Report {0}",
            "Recruitment Strategy {0}", "Compensation Benchmarking {0}", "Retention Analysis Report {0}",
            "Employee Satisfaction Survey {0}", "Talent Development Plan {0}", "Workforce Planning Guide {0}"
        )
        Paragraphs = @(
            "This employee performance review guide establishes the framework for evaluating team member contributions. The onboarding process has been redesigned to improve new hire retention rates.",
            "Our benefits enrollment program offers comprehensive healthcare coverage and retirement planning options. Employee training budgets have increased to support continuous professional development.",
            "The diversity and inclusion initiative has achieved measurable progress in representation across all levels. Recruitment strategies now incorporate blind resume screening and structured interviews.",
            "Compensation benchmarking against industry peers ensures competitive salary ranges. Our retention strategy focuses on career growth pathways and employee engagement programs.",
            "The talent acquisition team has streamlined the recruitment pipeline, reducing time-to-hire while maintaining quality standards. Performance management now includes quarterly check-ins and continuous feedback."
        )
    }
    "Marketing" = @{
        Keywords = @("campaign", "brand", "digital", "SEO", "content", "social media", "analytics", "engagement", "conversion", "audience")
        Phrase = "digital marketing strategy"
        Titles = @(
            "Digital Marketing Strategy {0}", "Brand Guidelines Update {0}", "SEO Performance Report Q{0}",
            "Social Media Campaign Plan {0}", "Content Calendar {0}", "Analytics Dashboard Review {0}",
            "Audience Engagement Report {0}", "Conversion Optimization Study {0}", "Campaign ROI Analysis {0}",
            "Brand Awareness Survey {0}", "Marketing Automation Setup {0}", "Influencer Partnership Guide {0}"
        )
        Paragraphs = @(
            "Our digital marketing strategy focuses on integrated campaigns across social media channels and search engines. Brand consistency is maintained through updated guidelines and approved content templates.",
            "SEO performance has improved significantly with optimized content and technical improvements. The analytics dashboard tracks key metrics including engagement rate, conversion rate, and audience growth.",
            "Social media campaigns generated strong audience engagement across all platforms. Content marketing efforts emphasize thought leadership and brand storytelling to build trust with target audiences.",
            "The conversion optimization study identified key friction points in the customer journey. A/B testing of landing pages resulted in a measurable improvement in lead generation quality.",
            "Brand awareness metrics show positive trends following the integrated campaign launch. Digital advertising spend has been optimized using data-driven audience targeting and programmatic buying strategies."
        )
    }
    "Engineering" = @{
        Keywords = @("architecture", "API", "deployment", "sprint", "microservices", "pipeline", "testing", "scalability", "infrastructure", "code review")
        Phrase = "system architecture design"
        Titles = @(
            "System Architecture Design v{0}", "API Reference Guide {0}", "Deployment Pipeline Setup {0}",
            "Sprint Retrospective Report {0}", "Microservices Migration Plan {0}", "Testing Strategy Document {0}",
            "Scalability Assessment {0}", "Infrastructure Review {0}", "Code Review Standards {0}",
            "DevOps Practices Guide {0}", "Performance Optimization Plan {0}", "Security Architecture Review {0}"
        )
        Paragraphs = @(
            "This system architecture design document outlines the microservices-based approach for the next platform iteration. API contracts are defined using OpenAPI specifications with automated testing pipelines.",
            "The deployment pipeline has been enhanced with automated testing, security scanning, and staged rollout capabilities. Sprint velocity metrics indicate consistent delivery against committed backlog items.",
            "Our microservices architecture enables independent scaling and deployment of service components. Code review processes ensure quality standards while maintaining development velocity.",
            "Infrastructure scalability testing confirms the platform handles projected load increases. The testing strategy includes unit, integration, performance, and chaos engineering methodologies.",
            "The DevOps pipeline automates build, test, and deployment workflows. Container orchestration using Kubernetes provides reliable service discovery and automatic scaling capabilities."
        )
    }
    "Legal" = @{
        Keywords = @("contract", "compliance", "NDA", "intellectual property", "regulation", "liability", "arbitration", "amendment", "jurisdiction", "indemnity")
        Phrase = "non-disclosure agreement"
        Titles = @(
            "Standard NDA Template v{0}", "Compliance Audit Report {0}", "Contract Amendment {0}",
            "Intellectual Property Policy {0}", "Regulatory Compliance Guide {0}", "Liability Assessment {0}",
            "Arbitration Procedures {0}", "Data Privacy Framework {0}", "Vendor Agreement Template {0}",
            "Employment Contract Standards {0}", "Jurisdiction Analysis {0}", "Indemnity Clause Review {0}"
        )
        Paragraphs = @(
            "This non-disclosure agreement template has been updated to reflect current regulatory requirements. All contract amendments must be reviewed by legal counsel before execution.",
            "The compliance audit identified areas requiring enhanced controls around data protection and intellectual property management. Regulatory changes in key jurisdictions necessitate policy updates.",
            "Our liability framework addresses indemnity obligations across vendor and customer agreements. Arbitration clauses have been standardized to reduce dispute resolution costs and timelines.",
            "Intellectual property protections have been strengthened through updated NDA templates and patent filing procedures. Contract management processes ensure compliance with applicable regulations.",
            "The data privacy framework aligns with global regulations including GDPR and CCPA requirements. All vendor agreements now include mandatory data processing addendums and security obligations."
        )
    }
    "Operations" = @{
        Keywords = @("supply chain", "logistics", "inventory", "procurement", "warehouse", "vendor", "SLA", "manufacturing", "quality control", "distribution")
        Phrase = "supply chain optimization"
        Titles = @(
            "Supply Chain Optimization Plan {0}", "Warehouse Operations Manual {0}", "Vendor SLA Template {0}",
            "Inventory Management Guide {0}", "Procurement Process Review {0}", "Logistics Performance Report {0}",
            "Quality Control Standards {0}", "Distribution Network Analysis {0}", "Manufacturing Efficiency {0}",
            "Facility Management Plan {0}", "SLA Compliance Report {0}", "Vendor Performance Scorecard {0}"
        )
        Paragraphs = @(
            "This supply chain optimization plan addresses key bottlenecks in procurement and distribution processes. Warehouse operations have been restructured to improve inventory turnover and reduce carrying costs.",
            "Vendor SLA compliance monitoring ensures service quality meets contractual requirements. Logistics performance metrics track delivery accuracy, lead times, and transportation cost efficiency.",
            "Quality control standards have been updated to incorporate industry best practices and customer feedback. The procurement process now includes automated approval workflows and spend analytics.",
            "Our distribution network analysis identified opportunities to reduce transit times and consolidate shipping routes. Inventory management improvements have reduced stockout incidents significantly.",
            "Manufacturing efficiency initiatives focus on lean principles and continuous improvement methodologies. Facility management plans ensure optimal utilization of warehouse and production space."
        )
    }
    "Sales" = @{
        Keywords = @("pipeline", "prospect", "commission", "territory", "CRM", "demo", "proposal", "deal", "quota", "account")
        Phrase = "sales pipeline review"
        Titles = @(
            "Sales Pipeline Review Q{0}", "Territory Analysis Report {0}", "CRM Integration Guide {0}",
            "Enterprise Proposal Template {0}", "Commission Structure {0}", "Account Management Plan {0}",
            "Demo Preparation Checklist {0}", "Quota Achievement Report {0}", "Deal Stage Analysis {0}",
            "Prospect Qualification Guide {0}", "Sales Training Manual {0}", "Customer Retention Strategy {0}"
        )
        Paragraphs = @(
            "The sales pipeline review reveals strong deal flow in the enterprise segment. Territory assignments have been optimized based on account potential and representative capacity.",
            "CRM data quality improvements enable more accurate pipeline forecasting and commission calculations. Our proposal templates have been updated to reflect current product capabilities and pricing.",
            "Account management strategies focus on expanding relationships within existing accounts. Demo preparation guidelines ensure consistent and compelling product presentations across all territories.",
            "Quota achievement tracking shows positive trends across most territories. The prospect qualification framework helps prioritize high-value opportunities and allocate resources effectively.",
            "The commission structure has been redesigned to incentivize strategic selling and account expansion. Sales training programs emphasize consultative selling techniques and competitive differentiation."
        )
    }
    "Projects" = @{
        Keywords = @("milestone", "deliverable", "stakeholder", "risk register", "Gantt", "sprint", "backlog", "scope", "resource allocation", "timeline")
        Phrase = "project status report"
        Titles = @(
            "Project Status Report {0}", "Resource Allocation Plan {0}", "Sprint Planning Guide {0}",
            "Risk Assessment Matrix {0}", "Stakeholder Communication Plan {0}", "Project Charter v{0}",
            "Milestone Tracking Report {0}", "Deliverable Acceptance Criteria {0}", "Scope Change Request {0}",
            "Timeline Review Document {0}", "Backlog Prioritization Guide {0}", "Lessons Learned Report {0}"
        )
        Paragraphs = @(
            "This project status report summarizes progress against key milestones and deliverables. The risk register has been updated to reflect newly identified risks and mitigation strategies.",
            "Resource allocation across active projects has been optimized to address capacity constraints. Sprint planning sessions use velocity-based forecasting for accurate timeline projections.",
            "Stakeholder communication ensures alignment on project scope and expected deliverables. The Gantt chart reflects current timeline adjustments and critical path dependencies.",
            "Backlog prioritization follows a value-weighted scoring model that considers business impact and technical complexity. Scope change requests require formal review and stakeholder approval.",
            "The lessons learned report captures insights from completed project phases to improve future execution. Resource allocation models now incorporate team capacity planning and skill-based assignments."
        )
    }
}

# ─── Helper Functions ─────────────────────────────────────────────────────────

function Get-WeightedChoice {
    param([string[]]$Items, [double[]]$CumulativeWeights)
    $r = Get-Random -Minimum 0.0 -Maximum 1.0
    for ($i = 0; $i -lt $CumulativeWeights.Count; $i++) {
        if ($r -lt $CumulativeWeights[$i]) { return $Items[$i] }
    }
    return $Items[$Items.Count - 1]
}

function Get-RandomUser {
    if ($script:ResolvedUsers.Count -eq 0) { return $null }
    return $script:ResolvedUsers[(Get-Random -Minimum 0 -Maximum $script:ResolvedUsers.Count)]
}

function Get-RandomDepartmentTerm {
    $dept = Get-WeightedChoice -Items $script:DeptNames -CumulativeWeights $script:DeptWeights
    $children = $script:DepartmentTerms[$dept]
    if ($children -and $children.Count -gt 0 -and (Get-Random -Minimum 0 -Maximum 3) -gt 0) {
        # 66% chance of child term, 33% chance of root term
        return $children[(Get-Random -Minimum 0 -Maximum $children.Count)]
    }
    return $dept
}

function Get-RandomTags {
    param([int]$Min = 1, [int]$Max = 3)
    $count = Get-Random -Minimum $Min -Maximum ($Max + 1)
    $shuffled = $script:TagTerms | Get-Random -Count $count
    return $shuffled
}

function Get-SpreadDate {
    param([int]$Index, [int]$Total, [int]$MonthSpan = 12)
    $daysBack = [Math]::Floor(($Index / [Math]::Max($Total, 1)) * $MonthSpan * 30)
    return (Get-Date).AddDays(-$daysBack)
}

function Get-RandomBudget {
    # Log-normal distribution: $1K to $5M
    $exp = 3 + (Get-Random -Minimum 0.0 -Maximum 3.7)
    return [Math]::Round([Math]::Pow(10, $exp), 2)
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$BaseDelayMs = 1000
    )
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            return (& $ScriptBlock)
        } catch {
            if ($attempt -eq $MaxRetries) { throw }
            $delay = $BaseDelayMs * [Math]::Pow(2, $attempt - 1)
            Write-Warning "  Attempt $attempt failed: $($_.Exception.Message). Retrying in $($delay)ms..."
            Start-Sleep -Milliseconds $delay
        }
    }
}

function Ensure-TermGroup {
    param([string]$Name)
    $group = Get-PnPTermGroup -Identity $Name -ErrorAction SilentlyContinue
    if ($group) {
        Write-Host "  [EXISTS] Term group '$Name'" -ForegroundColor Yellow
        return $group
    }
    Write-Host "  [CREATE] Term group '$Name'" -ForegroundColor Green
    $group = New-PnPTermGroup -Name $Name
    $script:stats.TermSetsCreated++
    return $group
}

function Ensure-TermSet {
    param([string]$Name, [string]$GroupName)
    $ts = Get-PnPTermSet -Identity $Name -TermGroup $GroupName -ErrorAction SilentlyContinue
    if ($ts) {
        Write-Host "  [EXISTS] Term set '$Name'" -ForegroundColor Yellow
        return $ts
    }
    Write-Host "  [CREATE] Term set '$Name'" -ForegroundColor Green
    $ts = New-PnPTermSet -Name $Name -TermGroup $GroupName -Lcid 1033
    $script:stats.TermSetsCreated++
    return $ts
}

function Ensure-Term {
    param([string]$Name, [string]$TermSetName, [string]$GroupName, [string]$ParentTermId = $null)
    $existing = Get-PnPTerm -Identity $Name -TermSet $TermSetName -TermGroup $GroupName -ErrorAction SilentlyContinue
    if ($existing) { return $existing }
    if ($ParentTermId) {
        $term = New-PnPTerm -Name $Name -TermSet $TermSetName -TermGroup $GroupName -ParentTerm $ParentTermId -Lcid 1033
    } else {
        $term = New-PnPTerm -Name $Name -TermSet $TermSetName -TermGroup $GroupName -Lcid 1033
    }
    $script:stats.TermsCreated++
    return $term
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
        Write-Host "  [EXISTS] Site column '$DisplayName' ($InternalName)" -ForegroundColor Yellow
        return $field
    }
    Write-Host "  [CREATE] Site column '$DisplayName' ($InternalName) [$Type]" -ForegroundColor Green
    $params = @{
        DisplayName  = $DisplayName
        InternalName = $InternalName
        Type         = $Type
        Group        = $Group
    }
    if ($Choices.Count -gt 0) {
        $params["Choices"] = $Choices
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
        Write-Host "  [EXISTS] Taxonomy column '$DisplayName' ($InternalName)" -ForegroundColor Yellow
        return $field
    }
    Write-Host "  [CREATE] Taxonomy column '$DisplayName' ($InternalName)" -ForegroundColor Green
    $params = @{
        DisplayName  = $DisplayName
        InternalName = $InternalName
        TermSetPath  = $TermSetPath
        Group        = $Group
    }
    if ($MultiValue) {
        $params["MultiValue"] = $true
    }
    $field = Add-PnPTaxonomyField @params
    $script:stats.SiteColumnsCreated++
    return $field
}

function Ensure-Library {
    param([string]$Name, [string]$Description, [string[]]$Columns)
    $list = Get-PnPList -Identity $Name -ErrorAction SilentlyContinue
    if ($list) {
        Write-Host "  [EXISTS] Library '$Name'" -ForegroundColor Yellow
    } else {
        Write-Host "  [CREATE] Library '$Name'" -ForegroundColor Green
        $list = New-PnPList -Title $Name -Template DocumentLibrary -EnableVersioning -OnQuickLaunch
        Set-PnPList -Identity $Name -Description $Description
        $script:stats.LibrariesCreated++
    }
    foreach ($col in $Columns) {
        $existingField = Get-PnPField -List $Name -Identity $col -ErrorAction SilentlyContinue
        if (-not $existingField) {
            try {
                Add-PnPField -List $Name -Field $col -ErrorAction Stop | Out-Null
            } catch {
                Write-Warning "  Could not add column '$col' to '$Name': $($_.Exception.Message)"
            }
        }
    }
    return $list
}

function Ensure-List {
    param([string]$Name, [string]$Description, [string[]]$SharedColumns, [array]$ExtraColumns)
    $list = Get-PnPList -Identity $Name -ErrorAction SilentlyContinue
    if ($list) {
        Write-Host "  [EXISTS] List '$Name'" -ForegroundColor Yellow
    } else {
        Write-Host "  [CREATE] List '$Name'" -ForegroundColor Green
        $list = New-PnPList -Title $Name -Template GenericList -EnableVersioning -OnQuickLaunch
        Set-PnPList -Identity $Name -Description $Description
        $script:stats.ListsCreated++
    }
    # Add shared site columns
    foreach ($col in $SharedColumns) {
        $existingField = Get-PnPField -List $Name -Identity $col -ErrorAction SilentlyContinue
        if (-not $existingField) {
            try {
                Add-PnPField -List $Name -Field $col -ErrorAction Stop | Out-Null
            } catch {
                Write-Warning "  Could not add shared column '$col' to '$Name': $($_.Exception.Message)"
            }
        }
    }
    # Add list-specific extra columns
    foreach ($colDef in $ExtraColumns) {
        $existingField = Get-PnPField -List $Name -Identity $colDef.Name -ErrorAction SilentlyContinue
        if ($existingField) { continue }
        try {
            $params = @{
                List         = $Name
                DisplayName  = $colDef.Name
                InternalName = $colDef.Name
                Type         = $colDef.Type
            }
            if ($colDef.Choices) {
                $params["Choices"] = $colDef.Choices
            }
            Add-PnPField @params -ErrorAction Stop | Out-Null
        } catch {
            Write-Warning "  Could not add column '$($colDef.Name)' to '$Name': $($_.Exception.Message)"
        }
    }
    return $list
}

function Ensure-FolderPath {
    param([string]$LibraryName, [string]$FolderPath)
    $segments = $FolderPath -split "/"
    $currentPath = $LibraryName
    foreach ($segment in $segments) {
        $targetPath = "$currentPath/$segment"
        try {
            $folder = Get-PnPFolder -Url $targetPath -ErrorAction SilentlyContinue
        } catch {
            $folder = $null
        }
        if (-not $folder -or $folder.Name -eq "") {
            try {
                Add-PnPFolder -Name $segment -Folder $currentPath -ErrorAction Stop | Out-Null
                $script:stats.FoldersCreated++
            } catch {
                # Folder may already exist despite Get failing
            }
        }
        $currentPath = $targetPath
    }
}

function Get-FolderPaths {
    param([int]$FolderDepth)
    switch ($FolderDepth) {
        3 { return @("2025/Q1", "2025/Q2", "2025/Q3", "2025/Q4", "2026/Q1", "Templates", "Archive") }
        2 { return @("Active", "Drafts", "Completed", "Templates") }
        1 { return @("Images", "Presentations", "Infographics") }
        default { return @() }
    }
}

# ─── OOXML Document Generation ────────────────────────────────────────────────

function New-DocxBytes {
    param([string]$Title, [string]$BodyText)
    $memStream = [System.IO.MemoryStream]::new()
    $archive = [System.IO.Compression.ZipArchive]::new($memStream, [System.IO.Compression.ZipArchiveMode]::Create, $true)

    # [Content_Types].xml
    $ct = $archive.CreateEntry("[Content_Types].xml")
    $w = [System.IO.StreamWriter]::new($ct.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    $w.Close()

    # _rels/.rels
    $rels = $archive.CreateEntry("_rels/.rels")
    $w = [System.IO.StreamWriter]::new($rels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    $w.Close()

    # word/_rels/document.xml.rels
    $drels = $archive.CreateEntry("word/_rels/document.xml.rels")
    $w = [System.IO.StreamWriter]::new($drels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    $w.Close()

    # word/document.xml
    $escTitle = [System.Security.SecurityElement]::Escape($Title)
    $escBody = [System.Security.SecurityElement]::Escape($BodyText)
    $doc = $archive.CreateEntry("word/document.xml")
    $w = [System.IO.StreamWriter]::new($doc.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>')
    $w.Write("<w:p><w:pPr><w:pStyle w:val=`"Heading1`"/></w:pPr><w:r><w:t>$escTitle</w:t></w:r></w:p>")
    # Split body into paragraphs
    $paragraphs = $escBody -split "`n"
    foreach ($p in $paragraphs) {
        $trimmed = $p.Trim()
        if ($trimmed) {
            $w.Write("<w:p><w:r><w:t>$trimmed</w:t></w:r></w:p>")
        }
    }
    $w.Write('</w:body></w:document>')
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

    # [Content_Types].xml
    $ct = $archive.CreateEntry("[Content_Types].xml")
    $w = [System.IO.StreamWriter]::new($ct.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
    $w.Close()

    # _rels/.rels
    $rels = $archive.CreateEntry("_rels/.rels")
    $w = [System.IO.StreamWriter]::new($rels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    $w.Close()

    # xl/_rels/workbook.xml.rels
    $wrels = $archive.CreateEntry("xl/_rels/workbook.xml.rels")
    $w = [System.IO.StreamWriter]::new($wrels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    $w.Close()

    # xl/workbook.xml
    $wbook = $archive.CreateEntry("xl/workbook.xml")
    $w = [System.IO.StreamWriter]::new($wbook.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>')
    $w.Close()

    # xl/styles.xml (minimal)
    $sty = $archive.CreateEntry("xl/styles.xml")
    $w = [System.IO.StreamWriter]::new($sty.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>')
    $w.Close()

    # Build shared strings and sheet data from body text
    $escTitle = [System.Security.SecurityElement]::Escape($Title)
    $words = $BodyText -split "\s+" | Select-Object -First 50
    $strings = @($escTitle) + ($words | ForEach-Object { [System.Security.SecurityElement]::Escape($_) })

    # xl/sharedStrings.xml
    $ss = $archive.CreateEntry("xl/sharedStrings.xml")
    $w = [System.IO.StreamWriter]::new($ss.Open())
    $w.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><sst xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" count=`"$($strings.Count)`" uniqueCount=`"$($strings.Count)`">")
    foreach ($s in $strings) {
        $w.Write("<si><t>$s</t></si>")
    }
    $w.Write('</sst>')
    $w.Close()

    # xl/worksheets/sheet1.xml
    $sheet = $archive.CreateEntry("xl/worksheets/sheet1.xml")
    $w = [System.IO.StreamWriter]::new($sheet.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
    # Title row
    $w.Write('<row r="1"><c r="A1" t="s"><v>0</v></c></row>')
    # Data rows (5 words per row)
    $rowIdx = 2
    for ($si = 1; $si -lt $strings.Count; $si += 5) {
        $w.Write("<row r=`"$rowIdx`">")
        $colLetters = @("A","B","C","D","E")
        for ($c = 0; $c -lt 5 -and ($si + $c) -lt $strings.Count; $c++) {
            $w.Write("<c r=`"$($colLetters[$c])$rowIdx`" t=`"s`"><v>$($si + $c)</v></c>")
        }
        $w.Write("</row>")
        $rowIdx++
    }
    $w.Write('</sheetData></worksheet>')
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

    $escTitle = [System.Security.SecurityElement]::Escape($Title)
    $escBody = [System.Security.SecurityElement]::Escape($BodyText.Substring(0, [Math]::Min($BodyText.Length, 500)))

    # [Content_Types].xml
    $ct = $archive.CreateEntry("[Content_Types].xml")
    $w = [System.IO.StreamWriter]::new($ct.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>')
    $w.Close()

    # _rels/.rels
    $rels = $archive.CreateEntry("_rels/.rels")
    $w = [System.IO.StreamWriter]::new($rels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
    $w.Close()

    # ppt/_rels/presentation.xml.rels
    $prels = $archive.CreateEntry("ppt/_rels/presentation.xml.rels")
    $w = [System.IO.StreamWriter]::new($prels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>')
    $w.Close()

    # ppt/presentation.xml
    $pres = $archive.CreateEntry("ppt/presentation.xml")
    $w = [System.IO.StreamWriter]::new($pres.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
    $w.Close()

    # ppt/slideMasters/_rels/slideMaster1.xml.rels
    $smrels = $archive.CreateEntry("ppt/slideMasters/_rels/slideMaster1.xml.rels")
    $w = [System.IO.StreamWriter]::new($smrels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
    $w.Close()

    # ppt/slideMasters/slideMaster1.xml
    $sm = $archive.CreateEntry("ppt/slideMasters/slideMaster1.xml")
    $w = [System.IO.StreamWriter]::new($sm.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>')
    $w.Close()

    # ppt/slideLayouts/_rels/slideLayout1.xml.rels
    $slrels = $archive.CreateEntry("ppt/slideLayouts/_rels/slideLayout1.xml.rels")
    $w = [System.IO.StreamWriter]::new($slrels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
    $w.Close()

    # ppt/slideLayouts/slideLayout1.xml
    $sl = $archive.CreateEntry("ppt/slideLayouts/slideLayout1.xml")
    $w = [System.IO.StreamWriter]::new($sl.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>')
    $w.Close()

    # ppt/slides/_rels/slide1.xml.rels
    $s1rels = $archive.CreateEntry("ppt/slides/_rels/slide1.xml.rels")
    $w = [System.IO.StreamWriter]::new($s1rels.Open())
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
    $w.Close()

    # ppt/slides/slide1.xml
    $s1 = $archive.CreateEntry("ppt/slides/slide1.xml")
    $w = [System.IO.StreamWriter]::new($s1.Open())
    $w.Write("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><p:sld xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`" xmlns:p=`"http://schemas.openxmlformats.org/presentationml/2006/main`"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=`"1`" name=`"`"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>")
    # Title shape
    $w.Write("<p:sp><p:nvSpPr><p:cNvPr id=`"2`" name=`"Title`"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=`"457200`" y=`"274638`"/><a:ext cx=`"8229600`" cy=`"1143000`"/></a:xfrm><a:prstGeom prst=`"rect`"/></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang=`"en-US`" sz=`"3200`" b=`"1`"/><a:t>$escTitle</a:t></a:r></a:p></p:txBody></p:sp>")
    # Body shape
    $w.Write("<p:sp><p:nvSpPr><p:cNvPr id=`"3`" name=`"Content`"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=`"457200`" y=`"1600200`"/><a:ext cx=`"8229600`" cy=`"4525963`"/></a:xfrm><a:prstGeom prst=`"rect`"/></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang=`"en-US`" sz=`"1800`"/><a:t>$escBody</a:t></a:r></a:p></p:txBody></p:sp>")
    $w.Write("</p:spTree></p:cSld></p:sld>")
    $w.Close()

    $archive.Dispose()
    $bytes = $memStream.ToArray()
    $memStream.Dispose()
    return $bytes
}

function New-PdfBytes {
    param([string]$Title, [string]$BodyText)
    # Minimal valid PDF with text content
    $safeTitle = $Title -replace '[()\\]', ' '
    $safeBody = ($BodyText -replace '[()\\]', ' ').Substring(0, [Math]::Min($BodyText.Length, 800))

    $streamContent = "BT /F1 16 Tf 72 750 Td ($safeTitle) Tj ET`nBT /F1 10 Tf 72 720 Td ($safeBody) Tj ET"
    $streamLength = [System.Text.Encoding]::ASCII.GetByteCount($streamContent)

    $pdf = @"
%PDF-1.4
1 0 obj <</Type /Catalog /Pages 2 0 R>> endobj
2 0 obj <</Type /Pages /Kids [3 0 R] /Count 1>> endobj
3 0 obj <</Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources <</Font <</F1 5 0 R>>>>>> endobj
4 0 obj <</Length $streamLength>>
stream
$streamContent
endstream endobj
5 0 obj <</Type /Font /Subtype /Type1 /BaseFont /Helvetica>> endobj
xref
0 6
0000000000 65535 f
0000000009 00000 n
0000000058 00000 n
0000000115 00000 n
0000000266 00000 n
0000000${($streamLength + 320).ToString().PadLeft(4, '0')} 00000 n
trailer <</Size 6 /Root 1 0 R>>
startxref
${($streamLength + 400)}
%%EOF
"@
    return [System.Text.Encoding]::ASCII.GetBytes($pdf)
}

function New-PlaceholderPngBytes {
    # Minimal 1x1 white PNG (67 bytes)
    return [byte[]]@(
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,  # PNG signature
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,  # IHDR chunk
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,  # 1x1 pixels
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
        0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,  # IDAT chunk
        0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
        0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
        0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,  # IEND chunk
        0x44, 0xAE, 0x42, 0x60, 0x82
    )
}

function Get-ThematicContent {
    param([string]$ThemeName, [int]$Variation, [string]$LibraryContext)
    $theme = $script:ThemeData[$ThemeName]
    if (-not $theme) { $theme = $script:ThemeData["Projects"] }

    $titleTemplates = $theme.Titles
    $paragraphs = $theme.Paragraphs
    $keywords = $theme.Keywords
    $phrase = $theme.Phrase

    # Select title template using variation
    $titleIdx = $Variation % $titleTemplates.Count
    $titleParam = if ($Variation % 4 -eq 0) { "2025" } elseif ($Variation % 4 -eq 1) { "2026" } elseif ($Variation % 4 -eq 2) { ($Variation % 4 + 1).ToString() } else { "v$($Variation % 5 + 1)" }
    $title = $titleTemplates[$titleIdx] -f $titleParam

    # Build body from 2-3 paragraphs + embedded phrase
    $paraCount = 2 + ($Variation % 2)
    $body = ""
    for ($p = 0; $p -lt $paraCount; $p++) {
        $pIdx = ($Variation + $p) % $paragraphs.Count
        $body += $paragraphs[$pIdx] + "`n`n"
    }
    # Embed the test phrase naturally
    if ($Variation % 3 -eq 0) {
        $body += "This document relates to our $phrase initiative for the $LibraryContext team.`n`n"
    }
    # Add some random keywords for variety
    $extraKeywords = $keywords | Get-Random -Count ([Math]::Min(4, $keywords.Count))
    $body += "Related topics: $($extraKeywords -join ', ')."

    return @{
        Title = $title
        Body  = $body
    }
}

function Get-FileExtension {
    param([int]$Index)
    $r = ($Index * 7 + 13) % 100  # Deterministic pseudo-random
    if ($r -lt 40) { return "docx" }
    if ($r -lt 60) { return "xlsx" }
    if ($r -lt 75) { return "pptx" }
    if ($r -lt 90) { return "pdf" }
    if ($r -lt 95) { return "txt" }
    return "png"
}

function New-DocumentBytes {
    param([string]$Extension, [string]$Title, [string]$Body)
    switch ($Extension) {
        "docx" { return New-DocxBytes -Title $Title -BodyText $Body }
        "xlsx" { return New-XlsxBytes -Title $Title -BodyText $Body }
        "pptx" { return New-PptxBytes -Title $Title -BodyText $Body }
        "pdf"  { return New-PdfBytes -Title $Title -BodyText $Body }
        "txt"  { return [System.Text.Encoding]::UTF8.GetBytes("$Title`n`n$Body") }
        "png"  { return New-PlaceholderPngBytes }
        default { return [System.Text.Encoding]::UTF8.GetBytes("$Title`n`n$Body") }
    }
}

function Get-SafeFileName {
    param([string]$Title, [string]$Extension)
    $safe = $Title -replace '[^\w\s\-]', '' -replace '\s+', '-'
    if ($safe.Length -gt 80) { $safe = $safe.Substring(0, 80) }
    return "$safe.$Extension"
}

# ─── Main Execution ──────────────────────────────────────────────────────────

try {
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  SP Search Test Data Provisioning" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "  Site:       $SiteUrl"
    Write-Host "  Docs/lib:   $DocumentsPerLibrary (total: $($DocumentsPerLibrary * 10))"
    Write-Host "  Items/list: $ItemsPerList (total: $($ItemsPerList * 10))"
    Write-Host ""

    # ═══ Phase 0: Validation & Connection ═══════════════════════════════════

    $step++
    Write-Host "[$step/$totalSteps] Validating prerequisites..." -ForegroundColor Cyan

    $pnpModule = Get-Module -ListAvailable -Name "PnP.PowerShell" | Select-Object -First 1
    if (-not $pnpModule) {
        Write-Error "PnP.PowerShell module not found. Install via: Install-Module PnP.PowerShell -Scope CurrentUser"
        exit 1
    }
    Write-Host "  PnP.PowerShell version: $($pnpModule.Version)" -ForegroundColor Gray
    Import-Module PnP.PowerShell -ErrorAction Stop

    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop

    # Connect
    Write-Host "  Connecting to SharePoint..." -ForegroundColor Gray
    if ($ClientId) {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
    } else {
        Connect-PnPOnline -Url $SiteUrl -Interactive
    }
    Write-Host "  Connected successfully" -ForegroundColor Green

    # Resolve users
    if ($Users.Count -eq 0) {
        $currentUser = Get-PnPProperty -ClientObject (Get-PnPWeb) -Property CurrentUser
        $script:ResolvedUsers = @($currentUser.LoginName)
        Write-Host "  Using current user: $($currentUser.LoginName)" -ForegroundColor Gray
    } else {
        $script:ResolvedUsers = $Users
        Write-Host "  Using $($Users.Count) provided users" -ForegroundColor Gray
    }

    # Clean existing if requested
    if ($CleanExisting) {
        Write-Warning "  Removing existing test libraries and lists..."
        foreach ($libDef in $script:LibraryDefs) {
            $existing = Get-PnPList -Identity $libDef.Name -ErrorAction SilentlyContinue
            if ($existing) {
                Remove-PnPList -Identity $libDef.Name -Force -ErrorAction SilentlyContinue
                Write-Host "  [REMOVED] Library '$($libDef.Name)'" -ForegroundColor Red
            }
        }
        foreach ($listDef in $script:ListDefs) {
            $existing = Get-PnPList -Identity $listDef.Name -ErrorAction SilentlyContinue
            if ($existing) {
                Remove-PnPList -Identity $listDef.Name -Force -ErrorAction SilentlyContinue
                Write-Host "  [REMOVED] List '$($listDef.Name)'" -ForegroundColor Red
            }
        }
    }

    # ═══ Phase 1: Term Store ════════════════════════════════════════════════

    $step++
    if ($SkipTermStore) {
        Write-Host "[$step/$totalSteps] Skipping term store provisioning (-SkipTermStore)" -ForegroundColor Yellow
    } else {
        Write-Host "[$step/$totalSteps] Provisioning term store..." -ForegroundColor Cyan

        $tg = Ensure-TermGroup -Name $TermGroupName

        # Departments (hierarchical)
        $deptTS = Ensure-TermSet -Name "Departments" -GroupName $TermGroupName
        foreach ($dept in $script:DepartmentTerms.Keys) {
            $parentTerm = Ensure-Term -Name $dept -TermSetName "Departments" -GroupName $TermGroupName
            foreach ($child in $script:DepartmentTerms[$dept]) {
                Ensure-Term -Name $child -TermSetName "Departments" -GroupName $TermGroupName -ParentTermId $parentTerm.Id | Out-Null
            }
        }

        # Document Types (flat)
        $docTypeTS = Ensure-TermSet -Name "Document Types" -GroupName $TermGroupName
        foreach ($dt in $script:DocumentTypeTerms) {
            Ensure-Term -Name $dt -TermSetName "Document Types" -GroupName $TermGroupName | Out-Null
        }

        # Regions (hierarchical)
        $regionTS = Ensure-TermSet -Name "Regions" -GroupName $TermGroupName
        foreach ($region in $script:RegionTerms.Keys) {
            $parentTerm = Ensure-Term -Name $region -TermSetName "Regions" -GroupName $TermGroupName
            foreach ($child in $script:RegionTerms[$region]) {
                Ensure-Term -Name $child -TermSetName "Regions" -GroupName $TermGroupName -ParentTermId $parentTerm.Id | Out-Null
            }
        }

        # Tags (flat)
        $tagTS = Ensure-TermSet -Name "Tags" -GroupName $TermGroupName
        foreach ($tag in $script:TagTerms) {
            Ensure-Term -Name $tag -TermSetName "Tags" -GroupName $TermGroupName | Out-Null
        }

        Write-Host "  Term store provisioning complete ($($script:stats.TermSetsCreated) sets, $($script:stats.TermsCreated) terms)" -ForegroundColor Green
    }

    # ═══ Phase 2: Site Columns ══════════════════════════════════════════════

    $step++
    Write-Host "[$step/$totalSteps] Creating site columns..." -ForegroundColor Cyan

    Ensure-SiteColumn -DisplayName "Status" -InternalName "SPSStatus" -Type "Choice" -Group $SiteColumnGroup -Choices $script:StatusChoices
    Ensure-SiteColumn -DisplayName "Priority" -InternalName "SPSPriority" -Type "Choice" -Group $SiteColumnGroup -Choices $script:PriorityChoices
    Ensure-SiteColumn -DisplayName "Region" -InternalName "SPSRegion" -Type "Choice" -Group $SiteColumnGroup -Choices $script:RegionChoices
    Ensure-SiteColumn -DisplayName "Budget" -InternalName "SPSBudget" -Type "Currency" -Group $SiteColumnGroup
    Ensure-SiteColumn -DisplayName "Is Active" -InternalName "SPSIsActive" -Type "Boolean" -Group $SiteColumnGroup
    Ensure-SiteColumn -DisplayName "Is Published" -InternalName "SPSIsPublished" -Type "Boolean" -Group $SiteColumnGroup
    Ensure-SiteColumn -DisplayName "Owner" -InternalName "SPSOwner" -Type "User" -Group $SiteColumnGroup
    Ensure-SiteColumn -DisplayName "Review Date" -InternalName "SPSReviewDate" -Type "DateTime" -Group $SiteColumnGroup
    Ensure-SiteColumn -DisplayName "External Ref" -InternalName "SPSExternalRef" -Type "Text" -Group $SiteColumnGroup
    Ensure-SiteColumn -DisplayName "Rating" -InternalName "SPSRating" -Type "Number" -Group $SiteColumnGroup
    Ensure-SiteColumn -DisplayName "View Count" -InternalName "SPSViewCount" -Type "Number" -Group $SiteColumnGroup

    # Taxonomy columns (only if term store was provisioned or exists)
    if (-not $SkipTermStore) {
        Ensure-TaxonomyColumn -DisplayName "Department" -InternalName "SPSDepartment" -Group $SiteColumnGroup -TermSetPath "$TermGroupName|Departments"
        Ensure-TaxonomyColumn -DisplayName "Document Type" -InternalName "SPSDocumentType" -Group $SiteColumnGroup -TermSetPath "$TermGroupName|Document Types"
        Ensure-TaxonomyColumn -DisplayName "Tags" -InternalName "SPSTags" -Group $SiteColumnGroup -TermSetPath "$TermGroupName|Tags" -MultiValue
    }

    Write-Host "  Site columns complete ($($script:stats.SiteColumnsCreated) created)" -ForegroundColor Green

    # ═══ Phase 3: Document Libraries ════════════════════════════════════════

    $step++
    Write-Host "[$step/$totalSteps] Creating document libraries..." -ForegroundColor Cyan

    foreach ($libDef in $script:LibraryDefs) {
        Ensure-Library -Name $libDef.Name -Description $libDef.Desc -Columns $libDef.Cols

        # Create folder structure
        $folders = Get-FolderPaths -FolderDepth $libDef.FolderDepth
        foreach ($fp in $folders) {
            Ensure-FolderPath -LibraryName $libDef.Name -FolderPath $fp
        }
    }

    Write-Host "  Libraries complete ($($script:stats.LibrariesCreated) created, $($script:stats.FoldersCreated) folders)" -ForegroundColor Green

    # ═══ Phase 4: Custom Lists ══════════════════════════════════════════════

    $step++
    Write-Host "[$step/$totalSteps] Creating custom lists..." -ForegroundColor Cyan

    foreach ($listDef in $script:ListDefs) {
        Ensure-List -Name $listDef.Name -Description $listDef.Desc -SharedColumns $listDef.SharedCols -ExtraColumns $listDef.ExtraCols
    }

    Write-Host "  Lists complete ($($script:stats.ListsCreated) created)" -ForegroundColor Green

    # ═══ Phase 5: Document Upload ═══════════════════════════════════════════

    $step++
    if ($SkipDocuments) {
        Write-Host "[$step/$totalSteps] Skipping document upload (-SkipDocuments)" -ForegroundColor Yellow
    } else {
        Write-Host "[$step/$totalSteps] Generating and uploading documents..." -ForegroundColor Cyan

        $totalDocs = $DocumentsPerLibrary * $script:LibraryDefs.Count
        $globalDocIdx = 0

        foreach ($libDef in $script:LibraryDefs) {
            $libName = $libDef.Name
            $themes = $libDef.Themes
            $folders = @("") + (Get-FolderPaths -FolderDepth $libDef.FolderDepth)  # Include root

            Write-Host "  Uploading to '$libName'..." -ForegroundColor Gray

            for ($i = 0; $i -lt $DocumentsPerLibrary; $i++) {
                $globalDocIdx++

                if ($globalDocIdx % 50 -eq 0) {
                    $pct = [Math]::Round(($globalDocIdx / $totalDocs) * 100)
                    Write-Progress -Activity "Uploading documents" -Status "$globalDocIdx of $totalDocs ($pct%)" -PercentComplete $pct
                }

                try {
                    # Select theme
                    $themeName = $themes[$i % $themes.Count]
                    $content = Get-ThematicContent -ThemeName $themeName -Variation $i -LibraryContext $libName

                    # Select file type
                    $ext = Get-FileExtension -Index ($globalDocIdx + $i)
                    $fileName = Get-SafeFileName -Title $content.Title -Extension $ext

                    # Generate document bytes
                    $bytes = New-DocumentBytes -Extension $ext -Title $content.Title -Body $content.Body

                    # Select folder
                    $folderPath = $folders[$i % $folders.Count]
                    $targetFolder = if ($folderPath) { "$libName/$folderPath" } else { $libName }

                    # Upload
                    $stream = [System.IO.MemoryStream]::new($bytes)
                    $file = Invoke-WithRetry -ScriptBlock {
                        Add-PnPFile -FileName $fileName -Folder $targetFolder -Stream $stream -ErrorAction Stop
                    }
                    $stream.Dispose()

                    # Set metadata
                    $metadataValues = @{}

                    if ($libDef.Cols -contains "SPSStatus") {
                        $metadataValues["SPSStatus"] = Get-WeightedChoice -Items $script:StatusChoices -CumulativeWeights $script:StatusWeights
                    }
                    if ($libDef.Cols -contains "SPSPriority") {
                        $metadataValues["SPSPriority"] = Get-WeightedChoice -Items $script:PriorityChoices -CumulativeWeights $script:PriorityWeights
                    }
                    if ($libDef.Cols -contains "SPSRegion") {
                        $metadataValues["SPSRegion"] = Get-WeightedChoice -Items $script:RegionChoices -CumulativeWeights $script:RegionWeights
                    }
                    if ($libDef.Cols -contains "SPSBudget") {
                        $metadataValues["SPSBudget"] = Get-RandomBudget
                    }
                    if ($libDef.Cols -contains "SPSIsActive") {
                        $metadataValues["SPSIsActive"] = ((Get-Random -Minimum 0 -Maximum 10) -lt 7)  # 70% true
                    }
                    if ($libDef.Cols -contains "SPSIsPublished") {
                        $metadataValues["SPSIsPublished"] = ((Get-Random -Minimum 0 -Maximum 10) -lt 6)  # 60% true
                    }
                    if ($libDef.Cols -contains "SPSOwner") {
                        $owner = Get-RandomUser
                        if ($owner) { $metadataValues["SPSOwner"] = $owner }
                    }
                    if ($libDef.Cols -contains "SPSReviewDate") {
                        $metadataValues["SPSReviewDate"] = (Get-Date).AddDays((Get-Random -Minimum 1 -Maximum 180))
                    }
                    if ($libDef.Cols -contains "SPSExternalRef") {
                        $metadataValues["SPSExternalRef"] = "REF-$($globalDocIdx.ToString('D5'))"
                    }
                    if ($libDef.Cols -contains "SPSRating") {
                        $metadataValues["SPSRating"] = [Math]::Round(1 + (Get-Random -Minimum 0.0 -Maximum 4.0), 1)
                    }
                    if ($libDef.Cols -contains "SPSViewCount") {
                        $metadataValues["SPSViewCount"] = Get-Random -Minimum 0 -Maximum 50000
                    }

                    if ($metadataValues.Count -gt 0) {
                        Set-PnPListItem -List $libName -Identity $file.ListItemAllFields.Id -Values $metadataValues -ErrorAction SilentlyContinue | Out-Null
                    }

                    $script:stats.DocumentsUploaded++

                    # Throttle protection
                    Start-Sleep -Milliseconds 50

                } catch {
                    $script:stats.Errors++
                    Write-Warning "  [ERROR] Failed to upload doc $i in '$libName': $($_.Exception.Message)"
                }
            }
        }
        Write-Progress -Activity "Uploading documents" -Completed
        Write-Host "  Documents complete ($($script:stats.DocumentsUploaded) uploaded, $($script:stats.Errors) errors)" -ForegroundColor Green
    }

    # ═══ Phase 6: List Item Creation ════════════════════════════════════════

    $step++
    if ($SkipListItems) {
        Write-Host "[$step/$totalSteps] Skipping list item creation (-SkipListItems)" -ForegroundColor Yellow
    } else {
        Write-Host "[$step/$totalSteps] Creating list items..." -ForegroundColor Cyan

        $totalItems = $ItemsPerList * $script:ListDefs.Count
        $globalItemIdx = 0

        # Name pools for list items
        $firstNames = @("James","Mary","John","Patricia","Robert","Jennifer","Michael","Linda","David","Elizabeth","William","Susan","Sarah","Thomas","Jessica","Daniel","Karen","Matthew","Nancy","Andrew")
        $lastNames = @("Smith","Johnson","Williams","Brown","Jones","Garcia","Miller","Davis","Rodriguez","Martinez","Anderson","Taylor","Thomas","Hernandez","Moore","Martin","Jackson","Thompson","White","Harris")
        $cities = @("New York","Los Angeles","Chicago","Houston","Phoenix","San Antonio","Dallas","San Jose","Austin","Jacksonville","London","Berlin","Paris","Tokyo","Sydney","Mumbai","Toronto","Sao Paulo","Buenos Aires","Bogota")
        $jobTitles = @("Software Engineer","Product Manager","Data Analyst","UX Designer","Project Manager","Business Analyst","QA Engineer","DevOps Engineer","Technical Writer","Marketing Manager","Sales Rep","HR Specialist","Finance Analyst","Operations Manager","Legal Counsel")

        foreach ($listDef in $script:ListDefs) {
            $listName = $listDef.Name
            Write-Host "  Populating '$listName'..." -ForegroundColor Gray

            for ($i = 0; $i -lt $ItemsPerList; $i++) {
                $globalItemIdx++

                if ($globalItemIdx % 50 -eq 0) {
                    $pct = [Math]::Round(($globalItemIdx / $totalItems) * 100)
                    Write-Progress -Activity "Creating list items" -Status "$globalItemIdx of $totalItems ($pct%)" -PercentComplete $pct
                }

                try {
                    $values = @{}

                    switch ($listName) {
                        "Projects" {
                            $values["Title"] = "Project $($firstNames[$i % $firstNames.Count]) $($i + 1)"
                            $values["SPSStatus"] = Get-WeightedChoice -Items $script:StatusChoices -CumulativeWeights $script:StatusWeights
                            $values["SPSPriority"] = Get-WeightedChoice -Items $script:PriorityChoices -CumulativeWeights $script:PriorityWeights
                            $values["SPSRegion"] = Get-WeightedChoice -Items $script:RegionChoices -CumulativeWeights $script:RegionWeights
                            $values["SPSBudget"] = Get-RandomBudget
                            $owner = Get-RandomUser; if ($owner) { $values["SPSOwner"] = $owner }
                            $values["StartDate"] = (Get-Date).AddDays(-(Get-Random -Minimum 30 -Maximum 365))
                            $values["EndDate"] = (Get-Date).AddDays((Get-Random -Minimum 30 -Maximum 365))
                            $values["PercentComplete"] = Get-Random -Minimum 0 -Maximum 101
                        }
                        "Contacts" {
                            $fn = $firstNames[$i % $firstNames.Count]
                            $ln = $lastNames[$i % $lastNames.Count]
                            $values["Title"] = "$fn $ln"
                            $values["Email"] = "$($fn.ToLower()).$($ln.ToLower())@contoso.com"
                            $values["Phone"] = "+1-555-$(Get-Random -Minimum 1000 -Maximum 9999)"
                            $values["JobTitle"] = $jobTitles[$i % $jobTitles.Count]
                            $values["SPSRegion"] = Get-WeightedChoice -Items $script:RegionChoices -CumulativeWeights $script:RegionWeights
                            $values["SPSIsActive"] = ((Get-Random -Minimum 0 -Maximum 10) -lt 8)
                        }
                        "Tasks" {
                            $values["Title"] = "Task: $($jobTitles[$i % $jobTitles.Count]) Item $($i + 1)"
                            $values["SPSStatus"] = Get-WeightedChoice -Items $script:StatusChoices -CumulativeWeights $script:StatusWeights
                            $values["SPSPriority"] = Get-WeightedChoice -Items $script:PriorityChoices -CumulativeWeights $script:PriorityWeights
                            $assignee = Get-RandomUser; if ($assignee) { $values["Assignee"] = $assignee }
                            $values["DueDate"] = (Get-Date).AddDays((Get-Random -Minimum -30 -Maximum 90))
                            $values["EstimatedHours"] = Get-Random -Minimum 1 -Maximum 80
                        }
                        "Events" {
                            $eventTypes = @("Team Meeting", "Sprint Review", "Quarterly Planning", "Training Session", "Product Demo", "Town Hall", "Workshop", "Hackathon")
                            $values["Title"] = "$($eventTypes[$i % $eventTypes.Count]) - $($cities[$i % $cities.Count])"
                            $values["EventDate"] = (Get-Date).AddDays((Get-Random -Minimum -60 -Maximum 120))
                            $values["EventEndDate"] = (Get-Date).AddDays((Get-Random -Minimum -59 -Maximum 121))
                            $values["Location"] = $cities[$i % $cities.Count]
                            $organizer = Get-RandomUser; if ($organizer) { $values["Organizer"] = $organizer }
                            $values["EventCategory"] = @("Meeting","Conference","Training","Social","Webinar")[$i % 5]
                        }
                        "Inventory" {
                            $itemNames = @("Laptop","Monitor","Keyboard","Mouse","Headset","Desk","Chair","Webcam","Docking Station","Printer","Scanner","Tablet","Phone","Projector","Whiteboard")
                            $values["Title"] = "$($itemNames[$i % $itemNames.Count]) - Unit $($i + 1)"
                            $values["Quantity"] = Get-Random -Minimum 1 -Maximum 500
                            $values["UnitPrice"] = [Math]::Round((Get-Random -Minimum 10.0 -Maximum 2500.0), 2)
                            $values["ItemCategory"] = @("Hardware","Software","Furniture","Supplies","Equipment")[$i % 5]
                            $values["Location"] = $cities[$i % $cities.Count]
                            $values["SPSIsActive"] = ((Get-Random -Minimum 0 -Maximum 10) -lt 8)
                        }
                        "Announcements" {
                            $announcementTopics = @("Company Update", "New Policy", "System Maintenance", "Team Achievement", "Holiday Schedule", "Office Relocation", "Benefits Change", "Training Opportunity")
                            $values["Title"] = "$($announcementTopics[$i % $announcementTopics.Count]) - $((Get-Date).AddDays(-$i).ToString('MMM yyyy'))"
                            $values["Body"] = "This announcement provides important information regarding $($announcementTopics[$i % $announcementTopics.Count].ToLower()). Please review the details and reach out to your manager with any questions."
                            $values["AnnouncementCategory"] = @("Company","Team","IT","HR","Facilities")[$i % 5]
                            $values["PublishDate"] = (Get-Date).AddDays(-(Get-Random -Minimum 0 -Maximum 180))
                            $values["ExpiryDate"] = (Get-Date).AddDays((Get-Random -Minimum 30 -Maximum 365))
                            $values["SPSPriority"] = Get-WeightedChoice -Items $script:PriorityChoices -CumulativeWeights $script:PriorityWeights
                        }
                        "Issues" {
                            $issueTypes = @("Login failure", "Data sync error", "Performance degradation", "UI rendering bug", "API timeout", "Permission denied", "Import failure", "Search not working", "Report generation error", "Notification delivery failure")
                            $values["Title"] = "$($issueTypes[$i % $issueTypes.Count]) #$($i + 1000)"
                            $values["Severity"] = @("Critical","Major","Minor","Trivial")[$i % 4]
                            $values["SPSStatus"] = Get-WeightedChoice -Items $script:StatusChoices -CumulativeWeights $script:StatusWeights
                            $assignee = Get-RandomUser; if ($assignee) { $values["Assignee"] = $assignee }
                            $values["ReportedDate"] = (Get-Date).AddDays(-(Get-Random -Minimum 1 -Maximum 90))
                            if ($i % 3 -eq 0) { $values["Resolution"] = "Issue resolved by applying the recommended configuration change and restarting the affected service." }
                        }
                        "FAQ" {
                            $questions = @("How do I reset my password?", "Where can I find the company handbook?", "How do I submit an expense report?", "What is the VPN setup process?", "How do I request time off?", "Where are the meeting room booking instructions?", "How do I access the training portal?", "What are the IT support hours?", "How do I update my contact information?", "Where can I find the org chart?")
                            $qIdx = $i % $questions.Count
                            $values["Title"] = "FAQ #$($i + 1)"
                            $values["Question"] = $questions[$qIdx]
                            $values["Answer"] = "Please follow the steps outlined in our internal knowledge base. For additional assistance, contact the relevant department via the help desk portal."
                            $values["FAQCategory"] = @("General","Technical","HR","Finance","IT")[$i % 5]
                            $values["HelpfulCount"] = Get-Random -Minimum 0 -Maximum 500
                        }
                        "Policies" {
                            $policyNames = @("Data Protection Policy", "Remote Work Policy", "Acceptable Use Policy", "Travel Expense Policy", "Code of Conduct", "Information Security Policy", "Social Media Policy", "Procurement Policy", "Health and Safety Policy", "Environmental Policy")
                            $values["Title"] = "$($policyNames[$i % $policyNames.Count]) v$([Math]::Ceiling(($i + 1) / $policyNames.Count))"
                            $values["PolicyNumber"] = "POL-$(($i + 1).ToString('D4'))"
                            $values["PolicyVersion"] = "$([Math]::Floor($i / 10 + 1)).$(($i % 10))"
                            $values["SPSReviewDate"] = (Get-Date).AddDays((Get-Random -Minimum 30 -Maximum 365))
                            $owner = Get-RandomUser; if ($owner) { $values["SPSOwner"] = $owner }
                            $values["SPSIsPublished"] = ((Get-Random -Minimum 0 -Maximum 10) -lt 6)
                        }
                        "Glossary" {
                            $terms = @("API", "SLA", "KPI", "ROI", "CAPEX", "OPEX", "MVP", "OKR", "CI/CD", "GDPR", "Microservices", "Agile", "Scrum", "DevOps", "Blockchain", "Machine Learning", "Cloud Native", "Zero Trust", "Kubernetes", "Terraform")
                            $tIdx = $i % $terms.Count
                            $values["Title"] = $terms[$tIdx]
                            $values["Definition"] = "A commonly used term in enterprise technology and business contexts. $($terms[$tIdx]) refers to a specific concept, methodology, or technology framework."
                            $values["GlossaryCategory"] = @("Technical","Business","Legal","Finance","HR")[$i % 5]
                            $values["RelatedTerms"] = "$($terms[($tIdx + 1) % $terms.Count]), $($terms[($tIdx + 2) % $terms.Count])"
                            $values["SPSIsActive"] = ((Get-Random -Minimum 0 -Maximum 10) -lt 9)
                        }
                    }

                    Invoke-WithRetry -ScriptBlock {
                        Add-PnPListItem -List $listName -Values $values -ErrorAction Stop | Out-Null
                    }
                    $script:stats.ListItemsCreated++

                    # Throttle
                    Start-Sleep -Milliseconds 30

                } catch {
                    $script:stats.Errors++
                    Write-Warning "  [ERROR] Failed to create item $i in '$listName': $($_.Exception.Message)"
                }
            }
        }
        Write-Progress -Activity "Creating list items" -Completed
        Write-Host "  List items complete ($($script:stats.ListItemsCreated) created)" -ForegroundColor Green
    }

    # ═══ Phase 7: Re-Index ══════════════════════════════════════════════════

    $step++
    if ($RequestReindex) {
        Write-Host "[$step/$totalSteps] Requesting site re-index..." -ForegroundColor Cyan
        try {
            Request-PnPReindexWeb -ErrorAction Stop
            Write-Host "  Site re-index requested. Crawl may take 15-60 minutes." -ForegroundColor Yellow
            Write-Host "  Monitor: SharePoint Admin Center > Search > Crawl Log" -ForegroundColor Gray
        } catch {
            Write-Warning "  Could not request re-index: $($_.Exception.Message)"
            Write-Host "  You can manually request via: Site Settings > Search and Offline Availability > Reindex site" -ForegroundColor Gray
        }
    } else {
        Write-Host "[$step/$totalSteps] Skipping re-index (use -RequestReindex to enable)" -ForegroundColor Yellow
    }

    # ═══ Phase 8: Mapping Guide ═════════════════════════════════════════════

    $step++
    Write-Host "[$step/$totalSteps] Managed Property Mapping Guide" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  After the search crawl completes, map these crawled properties" -ForegroundColor White
    Write-Host "  in SharePoint Admin Center > Search > Manage Search Schema:" -ForegroundColor White
    Write-Host ""
    Write-Host "  Crawled Property              | Map To               | Filter Type" -ForegroundColor White
    Write-Host "  ------------------------------|----------------------|------------------" -ForegroundColor Gray
    Write-Host "  ows_SPSStatus                 | RefinableString00    | Checkbox" -ForegroundColor Gray
    Write-Host "  ows_SPSPriority               | RefinableString01    | Checkbox" -ForegroundColor Gray
    Write-Host "  ows_SPSRegion                 | RefinableString02    | TagBox" -ForegroundColor Gray
    Write-Host "  ows_SPSDocumentType           | RefinableString03    | TagBox" -ForegroundColor Gray
    Write-Host "  ows_taxId_SPSDepartment       | RefinableString05    | Taxonomy tree" -ForegroundColor Gray
    Write-Host "  ows_taxId_SPSTags             | RefinableString06    | TagBox (multi)" -ForegroundColor Gray
    Write-Host "  ows_SPSIsActive               | RefinableString07    | Toggle" -ForegroundColor Gray
    Write-Host "  ows_SPSIsPublished            | RefinableString08    | Toggle" -ForegroundColor Gray
    Write-Host "  ows_SPSOwner                  | RefinableString09    | People" -ForegroundColor Gray
    Write-Host "  ows_SPSBudget                 | RefinableDecimal00   | Slider" -ForegroundColor Gray
    Write-Host "  ows_SPSRating                 | RefinableDecimal01   | Slider" -ForegroundColor Gray
    Write-Host "  ows_SPSViewCount              | RefinableInt00       | Sort/Number" -ForegroundColor Gray
    Write-Host "  ows_SPSReviewDate             | RefinableDate00      | DateRange" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Built-in (no mapping needed): Title, Author, Created, LastModifiedTime," -ForegroundColor DarkGray
    Write-Host "  FileType, Size, Path, contentclass, HitHighlightedSummary, etc." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  NOTE: Map AFTER first crawl completes, then request a second full crawl." -ForegroundColor Yellow

    # ═══ Summary ════════════════════════════════════════════════════════════

    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Green
    Write-Host "  Test Data Provisioning Complete!" -ForegroundColor Green
    Write-Host "======================================================" -ForegroundColor Green
    Write-Host "  Term sets created:    $($script:stats.TermSetsCreated)"
    Write-Host "  Terms created:        $($script:stats.TermsCreated)"
    Write-Host "  Site columns created: $($script:stats.SiteColumnsCreated)"
    Write-Host "  Libraries created:    $($script:stats.LibrariesCreated)"
    Write-Host "  Lists created:        $($script:stats.ListsCreated)"
    Write-Host "  Folders created:      $($script:stats.FoldersCreated)"
    Write-Host "  Documents uploaded:   $($script:stats.DocumentsUploaded)"
    Write-Host "  List items created:   $($script:stats.ListItemsCreated)"
    Write-Host "  Errors:               $($script:stats.Errors)"
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "======================================================" -ForegroundColor Red
    Write-Host "  Provisioning failed at Phase $step!" -ForegroundColor Red
    Write-Host "======================================================" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host "  Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    Write-Host "  Stack:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Partial results: $($script:stats.DocumentsUploaded) docs, $($script:stats.ListItemsCreated) items" -ForegroundColor Yellow
    Write-Host "  Re-run the script -- it is idempotent and will skip existing items." -ForegroundColor Yellow
    exit 1

} finally {
    try {
        $null = Get-PnPConnection -ErrorAction Stop
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host "  Disconnected from SharePoint" -ForegroundColor Gray
    } catch { }
}
